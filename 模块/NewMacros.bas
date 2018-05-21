Attribute VB_Name = "NewMacros"


Sub 全角转换()

    '利用数组转换有关内容：软回车转变为硬回车，全角数字、字母转换为半角数字、字母等
    Selection.WholeStory
    Dim C As Variant
    Dim D As Variant
    C = Array("^l", ",", ";", "０", "１", "２", "３", "４", "５", "６", "７", "８", "９", "．", "Ａ", "Ｂ", "Ｃ", "Ｄ", "Ｅ", "Ｆ", "Ｇ", _
        "Ｈ", "Ｉ", "Ｊ", "Ｋ", "Ｌ", "Ｍ", "Ｎ", "Ｏ", "Ｐ", "Ｑ", "Ｒ", "Ｓ", "Ｔ", "Ｕ", "Ｖ", "Ｗ", "Ｘ", "Ｙ", "Ｚ", _
        "ａ", "ｂ", "ｃ", "ｄ", "ｅ", "ｆ", "ｇ", "ｈ", "ｉ", "ｊ", "ｋ", "ｌ", "ｍ", "ｎ", "ｏ", "ｐ", "ｑ", "ｒ", "ｓ", _
        "ｔ", "ｕ", "ｖ", "ｗ", "ｘ", "ｙ", "ｚ")
    D = Array("^p", "，", "；", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
        "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", _
        "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")
    For i = 0 To 65
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = D(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    
End Sub



Sub 全角数字转换S()
'
' 全角数字转换S Macro
' 宏在 2003-6-27 由 DHG 录制
'
    Selection.WholeStory
    Dim A As Variant
    Dim B As Variant
    A = Array("０", "１", "２", "３", "４", "５", "６", "７", "８", "９")
    B = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    For i = 0 To 9
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = A(i)
        .Replacement.Text = B(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
End Sub
Sub 标题粗体C()
'
' 标题粗体C Macro
' 宏在 2003-6-27 由 DHG 录制
'
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
End Sub
Sub 拷贝文件中的表格并格式化T()

   Application.ScreenUpdating = False '关闭屏幕更新

'
' 拷贝文件中的表格并格式化T Macro
' 宏在 2003-6-27 由 DHG 录制
'
' 定义两个变量：NM -- 当前文档的文件名； T -- 文档的表格张数
    Dim NM As String
    Dim T As Integer
    
' 取得当前文件名
    Set MyDOC = Application.ActiveWindow.Document
'    MsgBox myDoc
    NM = Left(MyDOC, Len(MyDOC) - 4)
'    MsgBox NM
' 新建一个空白文档
    Documents.Add DocumentType:=wdNewBlankDocument
    Set MyDocN = Application.ActiveWindow.Document

' 把表格拷贝到新建文档中
    MyDOC.Activate
    T = ActiveDocument.Tables.Count ' 取得文档中表格的张数
'    MsgBox "这个文件共有 " & T & " 张表格"
For j = 1 To T
    ActiveDocument.Tables(j).Range.Copy
    MyDocN.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.Paste
    MyDOC.Activate
Next j

' 将新文档中的表格应用宏：表格B
    MyDocN.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
    ActiveDocument.Tables(j).Select
    Application.Run MacroName:="Normal.NewMacros.表格B"
Next j
' 给原文档中的表格加上名称(Tab1, Tab2, ...)，以便拷贝回来
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
Next j

' 删除原文档中的表格
For j = 1 To T
    If ActiveDocument.Tables.Count > 0 Then
    ActiveDocument.Tables(1).Range.Cut
    End If
Next j

'将宏 标准网页格式W 应用到文档中
Selection.WholeStory
Selection.Copy
Documents.Add DocumentType:=wdNewBlankDocument
Selection.Range.PasteSpecial DataType:=wdPasteText

    Application.Run MacroName:="Normal.NewMacros.标准网页格式W"
    Set myDocOK = Application.ActiveWindow.Document

For j = 1 To T
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
    MyDocN.Activate
    ActiveDocument.Tables(j).Range.Copy
    myDocOK.Activate
    Selection.Paste
    Next j
    myDocOK.Activate
    For Each Atable In ActiveDocument.Tables
        With Atable
             .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
             .Borders(wdBorderTop).LineWidth = wdLineWidth150pt
             .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
             .Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
        End With
    Next Atable
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & "_OK.doc")
    
       Application.ScreenUpdating = True '恢复屏幕更新

    
End Sub
Sub 删文档图片和表格并应用标准网页格式()

   Application.ScreenUpdating = False '关闭屏幕更新

'
' 定义三个变量：NM -- 文件名； N -- 图片数  T--表格数
    Dim NM As String
    Dim N As Integer
    Dim T As Integer
        
' 取得当前文件名
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    N = ActiveDocument.InlineShapes.Count ' 取得文档中图片数
    T = ActiveDocument.Tables.Count ' 取得文档中表格数

' 给原文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
    MyDOC.Activate
For j = 1 To N
    ActiveDocument.InlineShapes(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TU" & j
Next j
' 删除原文档中的图片
For j = 1 To N
    If ActiveDocument.InlineShapes.Count > 0 Then
    ActiveDocument.InlineShapes(1).Range.Cut
    End If
Next j

' 新建一个空白文档
    Documents.Add DocumentType:=wdNewBlankDocument
    Set MyDocN = Application.ActiveWindow.Document

' 把表格拷贝到新建文档中
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Range.Copy
    MyDocN.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.Paste
    MyDOC.Activate
Next j

' 将新文档中的表格应用宏：表格B
    MyDocN.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
    ActiveDocument.Tables(j).Select
    Application.Run MacroName:="Normal.NewMacros.表格B"
Next j
' 给原文档中的表格加上名称(Tab1, Tab2, ...)，以便拷贝回来
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
Next j

' 删除原文档中的表格
For j = 1 To T
    If ActiveDocument.Tables.Count > 0 Then
    ActiveDocument.Tables(1).Range.Cut
    End If
Next j

'将宏 A00_网页格式 应用到文档中
    A00_网页格式
    Set myDocOK = Application.ActiveWindow.Document
    myDocOK.Activate
    
    For j = 1 To T
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
    MyDocN.Activate
    ActiveDocument.Tables(j).Range.Copy
    myDocOK.Activate
    Selection.Paste
    Next j
    myDocOK.Activate
   
   Application.ScreenUpdating = True '恢复屏幕更新

'标准化后的文档另存,加了_OK
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & ".doc")
    ActiveDocument.SaveAs FileName:=NM & ".htm", FileFormat:=wdFormatHTML
    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & ".doc")
    Documents.Open (NM & "_OK.doc")
    MsgBox "文档已经按要求标准化，本文档共有： " & Chr(13) _
     & "    " & T & " 张表格" & Chr(13) _
     & "    " & N & " 张图片" & Chr(13) _
     & "图片保存在 " & NM & ".files 目录下"
    

    
    
End Sub
Sub 拷贝文档图片和表格到一个新文档()
   Application.ScreenUpdating = False '关闭屏幕更新

'
' 定义三个变量：NM -- 文件名； N -- 图片数  T--表格数
    Dim NM As String
    Dim N As Integer
    Dim T As Integer
    
    On Error Resume Next
        
' 取得当前文件名
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    N = ActiveDocument.InlineShapes.Count ' 取得文档中图片数
    T = ActiveDocument.Tables.Count ' 取得文档中表格数

' 给原文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
    MyDOC.Activate
For j = 1 To N
    ActiveDocument.InlineShapes(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TU" & j
Next j

' 新建一个空白文档
    Documents.Add DocumentType:=wdNewBlankDocument
    Set MyDocN = Application.ActiveWindow.Document

' 把表格拷贝到新建文档中
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Range.Copy
    MyDocN.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.Paste
    MyDOC.Activate
Next j

' 删除原文档中的图片
For j = 1 To N
    If ActiveDocument.InlineShapes.Count > 0 Then
    ActiveDocument.InlineShapes(1).Range.Cut
    MyDocN.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.Paste
    MyDOC.Activate

    End If
Next j

' 给原文档中的表格加上名称(Tab1, Tab2, ...)，以便拷贝回来
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
Next j

' 给新文档中的表格加上名称(Tab1, Tab2, ...)，以便拷贝回来
    MyDocN.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
Next j

' 给新文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
    MyDocN.Activate
For j = 1 To N
    ActiveDocument.InlineShapes(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TU" & j
Next j


' 删除原文档中的表格
    MyDOC.Activate
For j = 1 To T
    If ActiveDocument.Tables.Count > 0 Then
    ActiveDocument.Tables(1).Range.Cut
    End If
Next j

   Application.ScreenUpdating = True '恢复屏幕更新


'标准化后的文档另存,加了_OK
    ActiveDocument.SaveAs FileName:=NM & "_WZ.doc", FileFormat:=wdFormatDocument
    MyDocN.Activate
    ActiveDocument.SaveAs FileName:=NM & "_TB.doc", FileFormat:=wdFormatDocument
    MsgBox "文档已经按要求标准化，本文档共有： " & Chr(13) _
     & "    " & T & " 张表格" & Chr(13) _
     & "    " & N & " 张图片" & Chr(13)
    
End Sub

Sub A01_文图表()

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
    NM = Left(MyDOC, Len(MyDOC) - 5)
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
        ActiveDocument.SaveAs FileName:=NM & "_图片.doc", FileFormat:=wdFormatDocument
        
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
        ActiveDocument.SaveAs FileName:=NM & "_表格.doc", FileFormat:=wdFormatDocument
        
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
    ActiveDocument.SaveAs FileName:=NM & "_WZ.doc", FileFormat:=wdFormatDocument
    
    '-----------------------------------------------------
    
    '新建空白文档，把原文档中的文本拷贝至新文档，并执行“A00_网页格式”宏命令
    
    Documents.Add DocumentType:=wdNewBlankDocument
    Set DOC_OK = Application.ActiveWindow.Document
    
    页面设置 '执行“页面设置”宏命令
    
    MyDOC.Activate
    Selection.WholeStory
    Selection.Copy
    
    DOC_OK.Activate
    
    CommandBars("Office Clipboard").Visible = False
    Selection.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:= _
        wdInLine, DisplayAsIcon:=False
    
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
            .Execute FindText:="　　TU" & j
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
    'A01_批量处理表头

    '-----------------------------------------------------

   Application.ScreenUpdating = True '恢复屏幕更新


    DOC_OK.Activate

    '标准化后的文档另存,加了_OK
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument
    ' ActiveDocument.SaveAs FileName:=NM & "_TB.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & "_OK.doc")

    MsgBox "文档已经按要求标准化，本文档共有： " & Chr(13) _
     & "    " & TN & " 张表格" & Chr(13) _
     & "    " & CN & " 张图片" & Chr(13) _
     & "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒"
    
End Sub
Sub A01_文图表E()

' 定义三个变量：NM -- 文件名； n -- 图片数  t--表格数
    Dim NM As String
    Dim N As Integer
    Dim T As Integer

    st = VBA.Timer '程序运行计时器

   Application.ScreenUpdating = False '关闭屏幕更新
   On Error Resume Next '忽略错误
   
   A01_不合并表头表注E
        
' 取得当前文件名
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    N = ActiveDocument.InlineShapes.Count ' 取得文档中图片数
    T = ActiveDocument.Tables.Count ' 取得文档中表格数

' 给原文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
    MyDOC.Activate
For j = 1 To N
    ActiveDocument.InlineShapes(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TU" & j
Next j

' 新建一个空白文档
    Documents.Add DocumentType:=wdNewBlankDocument
    Set MyDocN = Application.ActiveWindow.Document
    网页页面

' 把表格拷贝到新建文档中
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Range.Copy
    MyDocN.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.Paste
    MyDOC.Activate
Next j

' 将原文档中的图片拷贝到新文档
For j = 1 To N
    If ActiveDocument.InlineShapes.Count > 0 Then
    ActiveDocument.InlineShapes(1).Range.Cut
    MyDocN.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.Paste
    MyDOC.Activate

    End If
Next j

' 给原文档中的表格加上名称(Tab1, Tab2, ...)，以便拷贝回来
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
Next j

' 给新文档中的表格加上名称(Tab1, Tab2, ...)，以便拷贝回来
    MyDocN.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
Next j

' 将新文档中的表格应用宏：表格E
    MyDocN.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
    ActiveDocument.Tables(j).Select
    表格E
Next j

' 给新文档中的图片加上名称(TU1, TU2, ...)，以便拷贝回来
    MyDocN.Activate
For j = 1 To N
    ActiveDocument.InlineShapes(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TU" & j
Next j


' 删除原文档中的表格
    MyDOC.Activate
For j = 1 To T
    If ActiveDocument.Tables.Count > 0 Then
    ActiveDocument.Tables(1).Range.Cut
    End If
Next j

'将宏 A01_英文网页格式 应用到文档中
    MyDOC.Activate
    
    网页页面
    A01_英文网页格式
    
    '把表格拷贝回来
    Set myDocOK = Application.ActiveWindow.Document
    myDocOK.Activate
    
    For j = 1 To T
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
    MyDocN.Activate
    ActiveDocument.Tables(j).Range.Copy
    myDocOK.Activate
    Selection.Paste
    Next j
    myDocOK.Activate
    
    Selection.HomeKey Unit:=wdStory
    
    '把图片拷贝回来
    For j = 1 To N
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
    MyDocN.Activate
    ActiveDocument.InlineShapes(1).Range.Cut
    myDocOK.Activate

    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceAtLeast
    Selection.ParagraphFormat.LineSpacing = 12
    Selection.Paste
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    'Selection.TypeParagraph
    
    Next j
    myDocOK.Activate
     Selection.HomeKey Unit:=wdStory
    空格替换
    A01_批量加粗表格合计行E
    A01_批量加粗表格中的特定行E
    A01_批量处理表头E
    
    Selection.HomeKey Unit:=wdStory
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p^p"
        .Replacement.Text = "^p^p"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.HomeKey Unit:=wdStory
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p^p"
        .Replacement.Text = "^p^p"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.EndKey Unit:=wdStory
    Selection.TypeBackspace
    Selection.HomeKey Unit:=wdStory
    
    A01_英文注释格式

   Application.ScreenUpdating = True '恢复屏幕更新


'标准化后的文档另存,加了_OK
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument
    MyDocN.Activate
   ' ActiveDocument.SaveAs FileName:=NM & "_TB.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & "_OK.doc")

    MsgBox "文档已经按要求标准化，本文档共有： " & Chr(13) _
     & "    " & T & " 张表格" & Chr(13) _
     & "    " & N & " 张图片" & Chr(13) _
     & "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒"
    
End Sub

Sub A01_英文注释格式()
    
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="Annotations:" & "^p"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
        End If
    End With
    
    Selection.Paragraphs(1).Range.Font.Bold = True
    Set MyRange = ActiveDocument.Range(start:=Selection.Range.start, End:=ActiveDocument.Range.End - 1)
    MyRange.Font.Italic = True
    Selection.HomeKey Unit:=wdStory

End Sub

Sub A01_标题加粗E()
    
    Selection.WholeStory
    Dim A As Variant
    A = Array("I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X" _
          , "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX")
    For j = 0 To 19
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="^p" & A(j) & "."
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
        Else
        Exit For
        End If
    End With
    Selection.Paragraphs(2).Range.Font.Bold = True
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next j

End Sub

Sub A01_英文网页格式()
    
    'st = VBA.Timer '程序运行计时器
    Application.ScreenUpdating = False '关闭屏幕更新
    Selection.HomeKey Unit:=wdStory

    Selection.WholeStory
    Selection.Cut
    Selection.Collapse Direction:=wdCollapseStart
    Selection.Range.PasteSpecial DataType:=wdPasteText
    
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
    
    Selection.WholeStory

    Dim O As Variant
    Dim R As Variant
    O = Array("^l", "  ", "^p^p", "^p", "　　^p")
    R = Array("^p", " ", "^p", "^p", "")
    
    For i = 0 To 4
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = O(i)
        .Replacement.Text = R(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    
    Selection.WholeStory

    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Paragraphs(1).Range.Select
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.WholeStory
    Selection.EndKey Unit:=wdStory
    Selection.TypeBackspace
    Selection.HomeKey Unit:=wdStory
    Eng_style
    A01_标题加粗E
    加空行
    
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p^p"
        .Replacement.Text = "^p^p"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    
    Selection.HomeKey Unit:=wdStory
        
    Application.ScreenUpdating = True '关闭屏幕更新
   ' MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒" '显示程序运行的时间


End Sub

Sub 调整分地区表格D()
'
' ALT+T
' 调整分地区表格D Macro
' 宏在 2003-5-23 由 DHG 录制
'
    Selection.InsertRowsAbove 1
    Selection.MoveDown Unit:=wdLine, Count:=2
    Selection.InsertRowsAbove 1
    Selection.MoveDown Unit:=wdLine, Count:=6
    Selection.InsertRowsAbove 1
    Selection.MoveDown Unit:=wdLine, Count:=4
    Selection.InsertRowsAbove 1
    Selection.MoveDown Unit:=wdLine, Count:=8
    Selection.InsertRowsAbove 1
    Selection.MoveDown Unit:=wdLine, Count:=7
    Selection.InsertRowsAbove 1
    Selection.MoveDown Unit:=wdLine, Count:=7
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.InsertRowsAbove 1
    Selection.MoveDown Unit:=wdLine, Count:=6
    Selection.InsertRowsAbove 1
    Selection.Tables(1).Select
    With Selection.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 12
        .WordWrap = True
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub

Sub 计算文档字数()

' 定义三个变量：NM -- 文件名； N -- 图片数  T--表格数 Z--字数
    Dim NM As String
    Dim N As Integer
    Dim T As Integer
    Dim Z As Integer
    Dim doc As Document
    Dim docFound As Boolean

    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)  ' 取得当前文件名
    N = ActiveDocument.InlineShapes.Count ' 取得文档中图片数
    T = ActiveDocument.Tables.Count ' 取得文档中表格数
    Z = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharacters) '字数
    D = Date
    MyDOC.Close SaveChanges:=wdDoNotSaveChanges

    For Each doc In Documents
        If InStr(1, doc.name, "稿费统计.doc", 1) Then
            doc.Activate
            docFound = True
            Exit For
        Else
            docFound = False
        End If
    Next doc

    If docFound = False Then FileName = "D:\00 word_dot"
    
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.InsertAfter Text:=D & Chr(9) & NM & Chr(9) & Z & Chr(9) & T & Chr(9) & N
    ActiveDocument.Save
    
End Sub
Sub MTest()

Set fs = Application.FileSearch
With fs
    .LookIn = "D:\02 MyDOC\02 统计制度\原始文件"
    .FileName = "*.DOC"
    If .Execute(SortBy:=msoSortByFileName, _
    SortOrder:=msoSortOrderAscending) > 0 Then
        For i = 1 To .FoundFiles.Count
    Documents.Open FileName:=.FoundFiles(i)
    Application.Run MacroName:="Normal.NewMacros.计算文档字数"

        Next i
    Else
        MsgBox "没找到可用文档"
    End If
End With
 
End Sub

Sub ShowFolderList()
    Dim fs, f, f1, s, sf
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder("F:\验证资料")
    Set sf = f.SubFolders
    For Each f1 In sf
        s = s & f1.name
        s = s & vbCrLf
    Next
    MsgBox s
End Sub

Sub ShowDriveList()
    Dim fs, dr, dc, s, N
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set dc = fs.Drives
    For Each dr In dc
        s = s & dr.DriveLetter & " - "
        If dr.DriveType = 3 Then
            N = dr.ShareName
      '  Else
      '      n = dr.VolumeName
        End If
        s = s & N & vbCrLf
    Next
    MsgBox s
End Sub

Sub 国统办()
    Selection.TypeText Text:="国家统计局办公室"
    Selection.TypeParagraph
    Selection.InsertDateTime DateTimeFormat:="EEEE年O月A日", InsertAsField:=False
    Selection.TypeParagraph
End Sub
Sub 日期1()
    Selection.InsertDateTime DateTimeFormat:="yyyy'年'M'月'd'日'", InsertAsField:=True
End Sub
Sub 日期2()
    Selection.InsertDateTime DateTimeFormat:="EEEE年O月A日", InsertAsField:=False
End Sub

Sub 回车替换()
    Application.ScreenUpdating = False '关闭屏幕更新

    '对选择的文本区域进行操作,如果没用选择,则自动选择整个文档
    Dim MyRange As Range
    If Len(Selection.Range.Text) = 0 Then Selection.WholeStory
    Set MyRange = Selection.Range
    
    With MyRang
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    End With
    Application.ScreenUpdating = True '关闭屏幕更新

End Sub

Sub 括号替换K()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "（"
        .Replacement.Text = "("
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "）"
        .Replacement.Text = ")"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub

Sub 作者单位()
    Selection.TypeText Text:="　　（作者单位：）"
    
End Sub

Sub 区划代码()
    Selection.WholeStory
    
        Dim A As Variant
A = Array("110000", "120000", "130000", "140000", "150000", "210000", "220000", "230000", "310000", "320000" _
          , "330000", "340000", "350000", "360000", "370000", "410000", "420000", "430000", "440000", "450000" _
          , "460000", "500000", "510000", "520000", "530000", "540000", "610000", "620000", "630000", "640000", _
          , "650000", "710000", "810000", "820000")
    
    For j = 0 To 34
    
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=A(j)
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
        Else
        Exit For
        End If
    End With
    Selection.Paragraphs(1).Range.Font.Bold = True
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeParagraph
    Next j

    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "  "
        .Replacement.Text = "　"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "　"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        Selection.MoveLeft Unit:=wdCharacter, Count:=1

    
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^l"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

 
End Sub

Sub 英文表M01()
    ActiveDocument.Tables(1).Cell(Row:=3, Column:=3).Select
    Selection.MoveDown Unit:=wdLine, Count:=7, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy

    ActiveDocument.Tables(2).Cell(Row:=4, Column:=2).Select
    Selection.MoveDown Unit:=wdLine, Count:=7, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Paste

    Set MyRange = Selection
    ActiveDocument.Tables(3).Cell(Row:=1, Column:=1).Select
    Selection.CopyFormat

    ActiveDocument.Tables(2).Cell(Row:=4, Column:=2).Select
    Selection.MoveDown Unit:=wdLine, Count:=7, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.PasteFormat

End Sub

Sub 英文表M02()
    ActiveDocument.Tables(1).Cell(Row:=3, Column:=2).Select
    Selection.MoveDown Unit:=wdLine, Count:=37, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy

    ActiveDocument.Tables(2).Cell(Row:=5, Column:=2).Select
    Selection.MoveDown Unit:=wdLine, Count:=37, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Paste

    Set MyRange = Selection
    ActiveDocument.Tables(3).Cell(Row:=1, Column:=1).Select
    Selection.CopyFormat

    ActiveDocument.Tables(2).Cell(Row:=5, Column:=2).Select
    Selection.MoveDown Unit:=wdLine, Count:=37, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.PasteFormat
    
    ActiveDocument.Tables(2).Cell(Row:=5, Column:=2).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Range.Font.Bold = True
    
End Sub

Sub 英文表M03()
    ActiveDocument.Tables(1).Cell(Row:=3, Column:=2).Select
    Selection.MoveDown Unit:=wdLine, Count:=37, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy

    ActiveDocument.Tables(2).Cell(Row:=5, Column:=2).Select
    Selection.MoveDown Unit:=wdLine, Count:=37, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Paste

    Set MyRange = Selection
    ActiveDocument.Tables(3).Cell(Row:=1, Column:=1).Select
    Selection.CopyFormat

    ActiveDocument.Tables(2).Cell(Row:=5, Column:=2).Select
    Selection.MoveDown Unit:=wdLine, Count:=37, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.PasteFormat
    
    ActiveDocument.Tables(2).Cell(Row:=5, Column:=2).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Range.Font.Bold = True
    
End Sub

Sub 英文表M04()
    ActiveDocument.Tables(1).Cell(Row:=3, Column:=3).Select
    Selection.MoveDown Unit:=wdLine, Count:=71, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Copy
    
    ActiveDocument.Tables(2).Cell(Row:=3, Column:=3).Select
    Selection.MoveDown Unit:=wdLine, Count:=71, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Paste
    
    Set MyRange = Selection
    ActiveDocument.Tables(3).Cell(Row:=1, Column:=1).Select
    Selection.CopyFormat

    ActiveDocument.Tables(2).Cell(Row:=3, Column:=3).Select
    Selection.MoveDown Unit:=wdLine, Count:=71, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.PasteFormat

End Sub


Sub 统计制度表头()

    ActiveDocument.Tables(1).Rows(1).Select
    For i = 1 To 3
        Selection.InsertRowsAbove 1
        Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
    Next i
    
    ActiveDocument.Tables(1).Cell(Row:=1, Column:=1).Select
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    
    Selection.Rows.HeightRule = wdRowHeightAuto
    Selection.Rows.Height = CentimetersToPoints(0)


    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone

    Options.DefaultBorderLineWidth = wdLineWidth150pt
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With


'_________________

A1 = ActiveDocument.Paragraphs(1).Range.Text
A2 = ActiveDocument.Paragraphs(2).Range.Text
A3 = ActiveDocument.Paragraphs(3).Range.Text
A4 = ActiveDocument.Paragraphs(4).Range.Text
A5 = ActiveDocument.Paragraphs(5).Range.Text
A6 = ActiveDocument.Paragraphs(6).Range.Text
A7 = ActiveDocument.Paragraphs(7).Range.Text
A8 = ActiveDocument.Paragraphs(8).Range.Text
MsgBox "The active document contains " & _
    ActiveDocument.Paragraphs.Count & " paragraphs."
Selection.TypeText Text:=A1

Selection.Collapse Direction:=wdCollapseEnd
If Selection.Information(wdWithInTable) = False Then
    Set mytable = _
        ActiveDocument.Tables.Add(Range:=Selection.Range, _
        NumRows:=3, NumColumns:=3)
    For Each aCell In mytable.Rows(1).Cells
        i = i + 1
        aCell.Range.Text = i
    Next aCell
End If

Selection.Tables(1).Rows(1).Select
Selection.Cells.Merge
Selection.Range.Text = A1
Selection.Tables(1).Cell(Row:=1, Column:=1).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace

Selection.Tables(1).Cell(Row:=3, Column:=1).Select
Selection.TypeText Text:=A6
Selection.Tables(1).Cell(Row:=3, Column:=1).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace

Selection.Tables(1).Cell(Row:=3, Column:=2).Select
Selection.TypeText Text:=A7
Selection.Tables(1).Cell(Row:=3, Column:=2).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace

Selection.Tables(1).Cell(Row:=3, Column:=3).Select
Selection.Range.Text = A8
Selection.Tables(1).Cell(Row:=3, Column:=3).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace

Selection.Tables(1).Cell(Row:=2, Column:=3).Select
Selection.TypeText Text:=A2 + A3 + A4 + A5
Selection.Tables(1).Cell(Row:=2, Column:=3).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
    
Selection.Tables(1).Select

    Application.Run MacroName:="回车替换"

End Sub

Sub Macro14()

i = 5
A = ActiveDocument.Paragraphs(1).Range.Text

Selection.Collapse Direction:=wdCollapseEnd
If Selection.Information(wdWithInTable) = False Then
    Set mytable = _
        ActiveDocument.Tables.Add(Range:=Selection.Range, _
        NumRows:=3, NumColumns:=5)
    For Each aCell In mytable.Rows(2).Cells
    
        i = i + 1
        aCell.Range.Text = i & A
    Next aCell
End If

End Sub

Sub 统计制度表格转换()
'
    Dim NM As String
    Dim T As Integer
    Dim str As String
    
    
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    Documents.Add DocumentType:=wdNewBlankDocument
    Set MyDocN = Application.ActiveWindow.Document

    MyDOC.Activate
    ActiveDocument.Tables(1).Range.Cut
    MyDocN.Activate
    Selection.EndKey Unit:=wdStory
'    Selection.TypeParagraph
    Selection.Paste
    MyDOC.Activate

    Selection.WholeStory
    Selection.Style = ActiveDocument.Styles("正文")
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 12
        .FirstLineIndent = CentimetersToPoints(0)
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .WordWrap = True
    End With
    With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 9
    End With
    
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="综合机关名称："
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeParagraph
    
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="计量单位"
                If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
            str = "计量单位"
        Else
        .Execute FindText:="有效期至"
            .Parent.Expand Unit:=wdParagraph
            str = "有效期至"
        End If
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=str
    End With
        
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeParagraph
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.Paragraphs(1).Range.Delete
    Selection.TypeText Text:="２００　年"
    Selection.TypeParagraph
    Selection.MoveDown Unit:=wdLine, Count:=2
    Selection.Paragraphs(1).Range.Delete
        Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="说明"
    End With
    Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    Selection.Cut

    MyDocN.Activate
    Selection.EndKey Unit:=wdStory
    Selection.Paste
    MyDOC.Activate
    Selection.TypeBackspace
    N = MyDOC.Paragraphs.Count
    
 ' MsgBox n

A1 = ActiveDocument.Paragraphs(1).Range.Text
A2 = ActiveDocument.Paragraphs(2).Range.Text
A3 = ActiveDocument.Paragraphs(3).Range.Text
A4 = ActiveDocument.Paragraphs(4).Range.Text
A5 = ActiveDocument.Paragraphs(5).Range.Text
A6 = ActiveDocument.Paragraphs(6).Range.Text
A7 = ActiveDocument.Paragraphs(7).Range.Text
A8 = ActiveDocument.Paragraphs(8).Range.Text

MyDocN.Activate
Selection.WholeStory
Selection.Cut
MyDOC.Activate
Selection.WholeStory
Selection.Paste

    MyDOC.Tables(1).Rows(1).Select
    For i = 1 To 3
        Selection.InsertRowsAbove 1
        Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
    Next i
    
    ActiveDocument.Tables(1).Cell(Row:=1, Column:=1).Select
    Selection.MoveDown Unit:=wdLine, Count:=2, Extend:=wdExtend
    Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
    
    Selection.Rows.HeightRule = wdRowHeightAuto
    Selection.Rows.Height = CentimetersToPoints(0)


    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone

    Options.DefaultBorderLineWidth = wdLineWidth150pt
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With

Selection.Tables(1).Rows(1).Select
Selection.Cells.Merge
Selection.Range.Text = A1
Selection.Tables(1).Cell(Row:=1, Column:=1).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
Selection.Tables(1).Cell(Row:=1, Column:=1).Select
    Application.Run MacroName:="表标题"

Selection.Tables(1).Cell(Row:=3, Column:=1).Select
Selection.TypeText Text:=A6
Selection.Tables(1).Cell(Row:=3, Column:=1).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
Selection.Tables(1).Cell(Row:=3, Column:=1).Select
    With Selection.ParagraphFormat
        .SpaceAfter = 3
        .Alignment = wdAlignParagraphJustify
    End With

Selection.Tables(1).Cell(Row:=3, Column:=2).Select
Selection.TypeText Text:=A7
Selection.Tables(1).Cell(Row:=3, Column:=2).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
Selection.Tables(1).Cell(Row:=3, Column:=2).Select
    With Selection.ParagraphFormat
        .SpaceAfter = 3
        .Alignment = wdAlignParagraphCenter
    End With


Selection.Tables(1).Cell(Row:=3, Column:=3).Select
Selection.Range.Text = A8
Selection.Tables(1).Cell(Row:=3, Column:=3).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
Selection.Tables(1).Cell(Row:=3, Column:=3).Select
    With Selection.ParagraphFormat
        .Alignment = wdAlignParagraphJustify
    End With


Selection.Tables(1).Cell(Row:=2, Column:=3).Select
Selection.TypeText Text:=A2 + A3 + A4 + A5
Selection.Tables(1).Cell(Row:=2, Column:=3).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
Selection.Tables(1).Cell(Row:=2, Column:=3).Select
    With Selection.ParagraphFormat
        .Alignment = wdAlignParagraphJustify
    End With

    
Selection.Tables(1).Select
    Application.Run MacroName:="回车替换"

Selection.Tables(1).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1

    Selection.InsertRowsBelow 1
    Selection.SelectRow
    Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
    
    Selection.TypeText Text:="单位负责人："
        Selection.SelectCell
    With Selection.ParagraphFormat
        .SpaceAfter = 3
        .Alignment = wdAlignParagraphJustify
    End With

    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="填表人："
           Selection.SelectCell
    With Selection.ParagraphFormat
        .SpaceAfter = 3
        .Alignment = wdAlignParagraphJustify
    End With
    
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="报出日期：２００　年　月　日"
               Selection.SelectCell
    With Selection.ParagraphFormat
        .SpaceAfter = 3
        .Alignment = wdAlignParagraphRight
    End With
    
        Selection.InsertRowsBelow 1
      Selection.SelectRow
    Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
    
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone

    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.SelectRow
    Selection.Cells.Merge
        
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.EndKey Unit:=wdStory, Extend:=wdExtend
    Selection.Cut
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.Paste
    
    Selection.SelectCell
    
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 12
        .Alignment = wdAlignParagraphJustify
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
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
    With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .name = ""
        .Size = 9
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
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
    Selection.TypeBackspace
    Selection.TypeBackspace
    
 Selection.Tables(1).Select
        Selection.Tables(1).Rows.LeftIndent = CentimetersToPoints(0)
    With Selection.Tables(1)
        .TopPadding = CentimetersToPoints(0)
        .BottomPadding = CentimetersToPoints(0)
        .LeftPadding = CentimetersToPoints(0)
        .RightPadding = CentimetersToPoints(0)
        .Spacing = 0
        .AllowPageBreaks = True
        .AllowAutoFit = False
    End With
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.5)
    
        With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 12
    End With



End Sub

Sub BT1()

Selection.Tables(1).Rows(1).Select
Selection.Cells.Merge
Selection.Range.Text = A1
Selection.Tables(1).Cell(Row:=1, Column:=1).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace

Selection.Tables(1).Cell(Row:=3, Column:=1).Select
Selection.TypeText Text:=A6
Selection.Tables(1).Cell(Row:=3, Column:=1).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace

Selection.Tables(1).Cell(Row:=3, Column:=2).Select
Selection.TypeText Text:=A7
Selection.Tables(1).Cell(Row:=3, Column:=2).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace

Selection.Tables(1).Cell(Row:=3, Column:=3).Select
Selection.Range.Text = A8
Selection.Tables(1).Cell(Row:=3, Column:=3).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace

Selection.Tables(1).Cell(Row:=2, Column:=3).Select
Selection.TypeText Text:=A2 + A3 + A4 + A5
Selection.Tables(1).Cell(Row:=2, Column:=3).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
    
Selection.Tables(1).Select

    Application.Run MacroName:="回车替换"

End Sub

Sub BT2()

Selection.Tables(1).Rows(1).Select
Selection.Cells.Merge
Selection.Range.Text = A1
Selection.Tables(1).Cell(Row:=1, Column:=1).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace

Selection.Tables(1).Cell(Row:=3, Column:=1).Select
Selection.TypeText Text:=A5
Selection.Tables(1).Cell(Row:=3, Column:=1).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace

Selection.Tables(1).Cell(Row:=3, Column:=2).Select
Selection.TypeText Text:=A6
Selection.Tables(1).Cell(Row:=3, Column:=2).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace

Selection.Tables(1).Cell(Row:=3, Column:=3).Select
Selection.Range.Text = A7
Selection.Tables(1).Cell(Row:=3, Column:=3).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace

Selection.Tables(1).Cell(Row:=2, Column:=3).Select
Selection.TypeText Text:=A2 + A3 + A4
Selection.Tables(1).Cell(Row:=2, Column:=3).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeBackspace
    
Selection.Tables(1).Select

    Application.Run MacroName:="回车替换"



End Sub

Sub 删除英文字母()

    Dim C As Variant
    Dim D As Variant
    C = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", _
              "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", " ", "　")
    For i = 0 To 53
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = ""
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    

End Sub

Sub 删标点符号()

    For i = 1 To 15
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(32 + i)
        .Replacement.Text = ""
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i

    For i = 1 To 7
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(57 + i)
        .Replacement.Text = ""
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i

    For i = 1 To 6
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(90 + i)
        .Replacement.Text = ""
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i

    For i = 1 To 4
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(122 + i)
        .Replacement.Text = ""
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i


End Sub

Sub 省份格式()

    Dim C As Variant
    Dim D As Variant
    C = Array("全国", "北京", "天津", "河北", "山西", "辽宁", "吉林", "上海", "江苏", "浙江", "安徽", "福建", "江西", "山东", _
    "河南", "湖北", "湖南", "广东", "广西", "海南", "重庆", "四川", "贵州", "云南", "西藏", "陕西", "甘肃", "青海", "宁夏", "新疆")
    D = Array("全　国", "北　京", "天　津", "河　北", "山　西", "辽　宁", "吉　林", "上　海", "江　苏", "浙　江", "安　徽", "福　建", "江　西", "山　东", _
    "河　南", "湖　北", "湖　南", "广　东", "广　西", "海　南", "重　庆", "四　川", "贵　州", "云　南", "西　藏", "陕　西", "甘　肃", "青　海", "宁　夏", "新　疆")
    
    For i = 0 To 29
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = D(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i

End Sub

Sub A01_省份替换()

    Dim C As Variant
    Dim D As Variant
    C = Array("总 计", "北 京", "天 津", "河 北", "山 西", "辽 宁", "吉 林", "上 海", "江 苏", "浙 江", "安 徽", "福 建", "江 西", "山 东", _
    "河 南", "湖 北", "湖 南", "广 东", "广 西", "海 南", "重 庆", "四 川", "贵 州", "云 南", "西 藏", "陕 西", "甘 肃", "青 海", "宁 夏", "新 疆")
    D = Array("总计", "北京", "天津", "河北", "山西", "辽宁", "吉林", "上海", "江苏", "浙江", "安徽", "福建", "江西", "山东", _
    "河南", "湖北", "湖南", "广东", "广西", "海南", "重庆", "四川", "贵州", "云南", "西藏", "陕西", "甘肃", "青海", "宁夏", "新疆")
    
    For i = 0 To 29
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = D(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i

End Sub
Sub 各地要事()

A00_网页格式

Application.ScreenUpdating = False '关闭屏幕更新


'对选择的文本区域进行操作,如果没用选择,则自动选择整个文档
Dim MyRange As Range
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If

    Set MyRange = Selection.Range
    
    '定义一个32个元素的数组,元素为各省区市
    Dim C As Variant
    C = Array("北京", "天津", "河北", "山西", "内蒙古", "辽宁", "吉林", "黑龙江", "上海", "江苏", "浙江", "安徽", "福建", "江西", "山东", _
    "河南", "湖北", "湖南", "广东", "广西", "海南", "重庆", "四川", "贵州", "云南", "西藏", "陕西", "甘肃", "青海", "宁夏", "新疆", "新疆生产建设兵团")
   
   '逐一查找并进行格式转换,对省份进行字体加粗,并相应增加空行
   For i = 0 To 31
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=C(i) & "^p"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
            
    '字体加粗
    Selection.Paragraphs(1).Range.Font.Bold = True
    OLD = Selection.Paragraphs(1).Range.Text
    
    '删除段前空格,段前和段后增加一空行
    Selection.Paragraphs(1).Range.Text = Chr(13) & Trim(OLD) & Chr(13)
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Else
    Selection.HomeKey Unit:=wdStory
        End If
    End With
    Next i
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Paragraphs(1).Range.Select
    Selection.Delete
 Application.ScreenUpdating = True '恢复屏幕更新

End Sub
Sub 页面设置()

    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
       ' .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
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
        .GutterPos = wdGutterPosLeft
        .LayoutMode = wdLayoutModeLineGrid
    End With
End Sub
Sub 缺省页面设置()

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
        .LeftMargin = CentimetersToPoints(3.17)
        .RightMargin = CentimetersToPoints(3.17)
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
        .GutterPos = wdGutterPosLeft
        .LayoutMode = wdLayoutModeLineGrid
    End With
End Sub

Sub 页面横()
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientLandscape
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
        .PageWidth = CentimetersToPoints(29.7)
        .PageHeight = CentimetersToPoints(21)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .GutterPos = wdGutterPosLeft
        .LayoutMode = wdLayoutModeLineGrid
    End With
End Sub

Sub Eng_style()

Set docOld = Application.ActiveWindow.Document
    Documents.Open FileName:="F:\00 Word_Dot\Demo-style_Eng.doc"
    ActiveDocument.Paragraphs(1).Range.Copy
    docOld.Activate
    Selection.HomeKey Unit:=wdStory
    Selection.Paste
    Documents("Demo-style_Eng.doc").Close SaveChanges:=wdDoNotSaveChanges

    docOld.Activate
ActiveDocument.Paragraphs(1).Range.Select
Selection.CopyFormat
Selection.WholeStory
Selection.PasteFormat
ActiveDocument.Paragraphs(1).Range.Delete


End Sub
Sub 文本X()
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
        .FirstLineIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0

    End With
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = "^p　　"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub tabletest()
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
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
     
Else
    MsgBox "【注意】插入点不在表格中！" & Chr(13) & _
           "请将插入点放到表格的任意单元格中， " & Chr(13) & _
           "然后再执行本宏，谢谢！"
   End If

End Sub


Sub Macro27()
    Set myDocOK = Application.ActiveWindow.Document
    Documents.Open ("D:\02 MyDOC\douhao.doc")
    Set dh = Application.ActiveWindow.Document
    dh.Characters(1).Copy
    
    myDocOK.Activate
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "y1"
        .Replacement.Text = "^c"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    dh.Activate
    dh.Characters(2).Copy
    
    myDocOK.Activate
    With Selection.Find
        .Text = "y2"
        .Replacement.Text = "^c"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    dh.Close
    myDocOK.Activate
End Sub

Sub 引号()
Selection.HomeKey Unit:=wdStory

'区分引号的第一个和第2个，分别标记为y1,y2
With ActiveDocument.Content.Find
    .ClearFormatting
    Do While .Execute(FindText:="""", Forward:=True, Format:=False) = True
        
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(34)
        .Replacement.Text = "y1"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
        
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(34)
        .Replacement.Text = "y2"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceOne

    .Forward = True
    .Wrap = wdFindContinue
    Loop
End With

'替换为正确的中文引号
    Set myDocOK = Application.ActiveWindow.Document
    Documents.Open ("D:\00 Word_Dot\yinhao.doc")
    Set dh = Application.ActiveWindow.Document
    dh.Characters(1).Copy
    
    myDocOK.Activate
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "y1"
        .Replacement.Text = "^c"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    dh.Activate
    dh.Characters(2).Copy
    
    myDocOK.Activate
    With Selection.Find
        .Text = "y2"
        .Replacement.Text = "^c"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    dh.Close
    myDocOK.Activate
    
    End Sub
Sub 删除数字()


    Dim C As Variant
    Dim D As Variant
    C = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    For i = 0 To 9
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = ""
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i

End Sub

Sub 粘贴文本()
    On Error Resume Next
    Selection.Range.PasteSpecial DataType:=wdPasteText
End Sub
Sub AutoEdit()
    Application.DisplayAlerts = wdAlertsNone

Set fs = Application.FileSearch
With fs
    .LookIn = "D:\test"
    .FileName = "*.htm"
    If .Execute(SortBy:=msoSortByFileName, _
    SortOrder:=msoSortOrderAscending) > 0 Then
        For i = 1 To .FoundFiles.Count
    Documents.Open FileName:=.FoundFiles(i)
    ActiveDocument.Tables(1).Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Delete Unit:=wdCharacter, Count:=1
    Dim NM As String
' 取得当前文件名
    Set MyDOC = Application.ActiveWindow.Document
'    MsgBox myDoc
    NM = Left(MyDOC, Len(MyDOC) - 4)
    ActiveDocument.SaveAs FileName:=NM & "1" & ".htm", FileFormat:=wdFormatHTML

  '      Application.DisplayAlerts = wdAlertsNone
  '  ActiveDocument.Close SaveChanges:=wdSaveChanges
     ActiveWindow.Close wdDoNotSaveChanges

        Next i
    Else
        MsgBox "没找到可用文档"
    End If
End With
 
End Sub

Sub 中文空格()
    Selection.TypeText Text:="　"
End Sub
Sub 表格转置()

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
Sub 翻译文稿1()
    Dim NM As String
    Dim P As Integer
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    Application.Run MacroName:="Normal.NewMacros.页面设置"

    pn = ActiveDocument.Content.Paragraphs.Count
    PT = 60
    If pn < PT Then
    Selection.WholeStory
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, Count:=1
    For j = 1 To PT - pn
    Selection.TypeParagraph
    Next j
    End If
    
    Dim PS(1 To 60) As Variant
    For i = 1 To pn
    PS(i) = Left(ActiveDocument.Paragraphs(i).Range.Text, Len(ActiveDocument.Paragraphs(i).Range.Text) - 1)
'    MsgBox PS(i)
    Next i
   
    Selection.WholeStory
    Selection.Cut
    Set mytable = ActiveDocument.Tables.Add(Range:=Selection.Range, _
    NumRows:=60, NumColumns:=2)
    Num = 1
    For Each aCell In ActiveDocument.Tables(1).Columns(1).Cells
    aCell.Range.Text = PS(Num)
        Num = Num + 1
    Next aCell
    
    Selection.WholeStory
    'Application.Run MacroName:="Normal.NewMacros.回车替换"

    ActiveDocument.Tables(1).Columns(1).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(1).PreferredWidth = 40
    ActiveDocument.Tables(1).Columns(2).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(2).PreferredWidth = 60
    Selection.HomeKey Unit:=wdStory
    'Application.Run MacroName:="Normal.NewMacros.删除空格K"

End Sub
Sub 文本加制表位()

    i = 1
    pn = ActiveDocument.Content.Paragraphs.Count
    For i = 1 To pn - 1

Dim T, C, P
C = "|"
T = Trim(ActiveDocument.Paragraphs(i).Range.Text)
L = Len(T)
P = InStr(1, T, C, 1)
S1 = Trim(Left(T, P - 1))
S2 = Trim(Right(T, L - P))
' MsgBox "S1: " & S1
 
' Selection.TypeText Text:=S1 & Chr(13)
 'Selection.TypeText Text:=S2

 P1 = InStr(1, S2, " ", 1)
 Do Until P1 = 0
 If InStr(1, S2, " ", 1) > 0 Then
 P1 = InStr(1, S2, " ", 1)
 S2 = Trim(Left(S2, P1 - 1)) & Chr(9) & Trim(Right(S2, Len(S2) - P1))
End If
P1 = InStr(1, S2, " ", 1)
Loop

Selection.TypeText Text:=S1 & Chr(9) & S2

 Next i
 
End Sub

Sub 转换英文表格()
    Dim MyRange As Range
    Set MyRange = Selection.Range
    
    With MyRange
        pn = MyRange.Paragraphs.Count
        For i = 1 To pn
            T = Trim(MyRange.Paragraphs(i).Range.Text)
            L = Len(T)

        '把"|"替换为制表位 Chr(9)
        C = "|"
        T = Trim(MyRange.Paragraphs(i).Range.Text)
        L = Len(T)
        P = InStr(1, T, C, 1)
        S1 = Trim(Left(T, P - 1))
        S2 = Trim(Right(T, L - P))

        '判断大写字母，并删除大写字母前的内容
        For j = 1 To 26
        P1 = InStr(1, S1, Chr(j + 64), 0)
        If P1 > 0 Then
        S1 = Right(S1, Len(S1) - (P1 - 1))
        Exit For
        End If
        Next j

        p2 = InStr(1, S2, " ", 1)
        Do Until p2 = 0
        If InStr(1, S2, " ", 1) > 0 Then
        p2 = InStr(1, S2, " ", 1)
        S2 = Trim(Left(S2, p2 - 1)) & Chr(9) & Trim(Right(S2, Len(S2) - p2))
        End If
        p2 = InStr(1, S2, " ", 1)
        Loop
        
        MyRange.Paragraphs(i).Range.Text = S1 & Chr(9) & S2
        'Selection.TypeText Text:=S1 & Chr(9) & S2
    Next i
    End With

End Sub
Sub 区分第一列1()
    Dim MyRange As Range
    Set MyRange = Selection.Range
    Set MyRange1 = ActiveDocument _
    .Range(start:=ActiveDocument.Content.End - 1)

    
    With MyRange
        pn = MyRange.Paragraphs.Count
        For i = 1 To pn

        '把"|"替换为制表位 Chr(9)
        C = "|"
        T = MyRange.Paragraphs(i).Range.Text
        L = Len(T)
        P = InStr(1, T, C, 1)
        S1 = Left(T, P - 1)
        S2 = Right(T, L - P)
        
        MyRange1.InsertAfter S1 & Chr(9) & S2
        'Selection.TypeText Text:=S1 & Chr(9) & S2
    Next i
    End With

End Sub

Sub A01_删汉字()

    i = 1
    pn = ActiveDocument.Content.Paragraphs.Count
    
    For i = 1 To pn
        T = Trim(ActiveDocument.Paragraphs(i).Range.Text)
        L = Len(T)
        s = ""
        For j = 1 To L - 1
            ZF = Left(T, 1)
            If Abs(Asc(ZF)) < 255 Then
                s = s & Left(T, 1)
            End If
            T = Right(T, Len(T) - 1)
        Next j
        ActiveDocument.Paragraphs(i).Range.Text = s & Chr(13)
    Next i
    
End Sub
Sub 转换数字部分1()
    
    Dim MyRange As Range
    Dim PS() As Variant
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If

    Set MyRange = Selection.Range
    Set MyRange1 = ActiveDocument _
    .Range(start:=ActiveDocument.Content.End - 1)

    
    With MyRange
        pn = MyRange.Paragraphs.Count
        'For i = 1 To PN - 1

        '把"|"替换为制表位 Chr(9)
       ' C = "|"
        T = LTrim(MyRange.Paragraphs(1).Range.Text)
        L = Len(T)
       ' P = InStr(1, T, C, 1)
       ' S1 = Trim(Left(T, P - 1))
       ' S2 = Right(T, L - P)
        ReDim Preserve PS(1)
        PS(0) = 1


        S2 = LTrim(T)
        p2 = InStr(1, S2, " ", 1)
        
        Do Until p2 = 0
        If InStr(1, S2, " ", 1) > 0 Then
        p2 = InStr(1, S2, " ", 1)
        S3 = Trim(Left(S2, p2 - 1))
        S4 = Right(S2, Len(S2) - p2)
        Pi = InStr(1, T, S4, 1)
        L3 = Len(S3)
        L4 = Len(S4)
        
        ReDim Preserve PS(UBound(PS) + 1)
        PS(UBound(PS) - 1) = Pi - 1
        'MsgBox PS(UBound(PS) - 1)
        'MsgBox S1
        
        End If
        S2 = Trim(S4)
        p2 = InStr(1, S2, " ", 1)
        
        Loop
       ' Selection.TypeText Text:=S1 & Chr(9) & S2

       'MsgBox S1 & Chr(9) & S2
       ' MyRange.Paragraphs(i).Range.Text = S1 & Chr(9) & S2
        'Selection.TypeText Text:=S1 & Chr(9) & S2
    'Next i
        'Selection.TypeText Text:=S1 & Chr(9) & S2
    End With
    
    If PS(UBound(PS)) = 0 Then
        PS(UBound(PS)) = Len(T)
    End If
         'MsgBox UBound(PS)

    With MyRange
    For Each MyPara In MyRange.Paragraphs
        T = Trim(MyPara.Range.Text)
        If Len(T) = 1 Then
        Exit For
        End If
        S1 = ""
        s = S1 & Chr(9)
        For j = 1 To UBound(PS)
            ST1 = Right(T, Len(T) - PS(j - 1) + 1)
            ST2 = Left(ST1, PS(j) - PS(j - 1))
            s = s & Trim(ST2) & Chr(9)
        Next j
        With MyRange1
        .InsertAfter Left(s, Len(s) - 1) & Chr(13)
        End With
     Next MyPara
     End With

End Sub

Sub NewTableStyle()
    Dim styTable As Style

    Set styTable = ActiveDocument.Styles.Add( _
        name:="TableStyle 1", Type:=wdStyleTypeTable)

    With styTable.Table

        'Apply borders around table, a double border to the heading row,
        'a double border to the last column, and shading to last row
        .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        .Borders(wdBorderRight).LineStyle = wdLineStyleSingle

        .Condition(wdFirstRow).Borders(wdBorderBottom) _
            .LineStyle = wdLineStyleDouble

        .Condition(wdLastColumn).Borders(wdBorderLeft) _
            .LineStyle = wdLineStyleDouble

        .Condition(wdLastRow).Shading _
            .BackgroundPatternColor = wdColorGray125

    End With
    
'
For Each MyPara In ActiveDocument.Paragraphs
    MyPara.Range.Select
        With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 12
        .Bold = False
        
    End With
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
    End With

Next MyPara

'

End Sub

Sub 所有表格的顶线和底线加粗()

For Each Atable In ActiveDocument.Tables
'    atable.Borders.OutsideLineStyle = wdLineStyleSingle
'    atable.Borders.OutsideLineWidth = wdLineWidth150pt
'    atable.Borders.InsideLineStyle = wdLineStyleNone
    With Atable
         .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
         .Borders(wdBorderTop).LineWidth = wdLineWidth150pt
         .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
         .Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
         
       ' .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
       ' .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
      '  .Condition(wdFirstRow).Borders(wdBorderTop).LineWidth = wdLineWidth150pt
      '  .Condition(wdLastRow).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
       ' .Condition(wdLastColumn).Borders(wdBorderLeft).LineStyle = wdLineStyleDouble
       ' .Condition(wdLastRow).Shading.BackgroundPatternColor = wdColorGray125

    End With
Next Atable

End Sub
Sub 表标题()
    
    T = Selection.Range.Text
    If Len(T) = 0 Then
    MsgBox "请先选择标题文本"
    Else
    Selection.Delete
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=2, NumColumns:= _
        1, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    Selection.Tables(1).Select
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        With Selection.Tables(1)
         .TopPadding = CentimetersToPoints(0)
         .BottomPadding = CentimetersToPoints(0)
         .LeftPadding = CentimetersToPoints(0)
         .RightPadding = CentimetersToPoints(0)
         .Spacing = 0
         .AllowPageBreaks = True
         .AllowAutoFit = True
        End With

    ActiveDocument.Tables(1).Columns(1).Cells(1).Range.Text = Left(T, Len(T) - 1)
    ActiveDocument.Tables(1).Columns(1).Cells(1).Select
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .name = "Times New Roman"
        .Size = 12
        .Bold = True
    End With
    End If
End Sub

Sub 区分第一列()
    Dim MyRange As Range
    Set MyRange = Selection.Range
    Set MyRange1 = ActiveDocument.Range(start:=ActiveDocument.Content.End - 1)
    
    With MyRange
        pn = MyRange.Paragraphs.Count
        For i = 1 To pn

        '把"|"替换为制表位 Chr(9)
        C = "|"
        T = MyRange.Paragraphs(i).Range.Text
        L = Len(T)
        P = InStr(1, T, C, 1)
        S1 = Left(T, P - 1)
        S2 = Right(T, L - P)
        
        MyRange.Paragraphs(i).Range.Text = Trim(S1) & Chr(9) & S2
        
        'MyRange1.InsertAfter Trim(S1) & Chr(9) & S2
        'Selection.TypeText Text:=S1 & Chr(9) & S2
    Next i
    End With
    
    With MyRange
    MyRange.Select
        Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=2, _
        NumRows:=pn, AutoFitBehavior:=wdAutoFitWindow
    End With


End Sub

Sub 转换主栏和数字部分()
    
    Dim MyRange As Range
    Dim NM As String
    'Dim T As Integer

    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    Set MyRange = Selection.Range
    MyRange.Copy
    
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    Documents.Add DocumentType:=wdNewBlankDocument
    Set MyDocN = Application.ActiveWindow.Document
    Selection.Paste
    Selection.TypeBackspace

    Selection.WholeStory
    Set MyRange1 = Selection.Range

    With MyRange1
        pn = MyRange1.Paragraphs.Count
        For i = 1 To pn
        '把"|"替换为制表位 Chr(9)
        C = "|"
        T = MyRange1.Paragraphs(i).Range.Text
        L = Len(T)
        P = InStr(1, T, C, 1)
        S1 = Left(T, P - 1)
        S2 = Right(T, L - P)
        MyRange1.Paragraphs(i).Range.Text = Trim(S1) & Chr(9) & S2
        Next i
    End With
    
    With MyRange1
    MyRange1.Select
        Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=2, _
        NumRows:=pn, AutoFitBehavior:=wdAutoFitWindow
    End With
    
    ActiveDocument.Content.Tables(1).Columns(2).Select
    Selection.Copy
    
    Documents.Add DocumentType:=wdNewBlankDocument
    Selection.TypeParagraph
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    Set myDocN1 = Application.ActiveWindow.Document
        Application.Run MacroName:="Normal.NewMacros.粘贴文本"
        PN1 = ActiveDocument.Content.Paragraphs.Count
            Set MyRange2 = ActiveDocument _
    .Range(start:=0, End:=ActiveDocument.Content.Paragraphs(PN1 - 1).Range.End)
    MyRange2.Select
        Application.Run MacroName:="Normal.NewMacros.转换数字部分"
    ActiveDocument.Content.Tables(1).Range.Select
    Selection.Copy
    ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
    MyDocN.Activate
        Selection.WholeStory
        Selection.EndKey Unit:=wdStory
        Selection.TypeParagraph
        Selection.Paste
        
        C1 = ActiveDocument.Content.Tables(1).Columns(1).Cells.Count
        C2 = ActiveDocument.Content.Tables(2).Columns(1).Cells.Count
        If C1 = C2 Then
        For i = 1 To C1
        TC1 = ActiveDocument.Content.Tables(1).Columns(1).Cells(i).Range.Text
        TC2 = Left(TC1, Len(TC1) - 1)
        ActiveDocument.Content.Tables(2).Columns(1).Cells(i).Range.Delete
        ActiveDocument.Content.Tables(2).Columns(1).Cells(i).Range.Text = TC2
        Next i
        End If
        ActiveDocument.Content.Tables(1).Range.Select
        Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
        Selection.Delete
        ActiveDocument.Content.Tables(1).Columns(1).Select
        Application.Run MacroName:="Normal.NewMacros.回车替换"
        ActiveDocument.Content.Tables(1).Range.Select
        Selection.Copy
        ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
        MyDOC.Activate
        Selection.Delete
        Selection.Paste
        ActiveDocument.Content.Tables(1).Columns(1).Cells(1).Select
        Selection.InsertRowsAbove 2

End Sub
Sub 转换数字部分()
    
    Dim MyRange As Range
    Dim PS() As Variant
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If

    Set MyRange = Selection.Range
    Set MyRange1 = ActiveDocument _
    .Range(start:=ActiveDocument.Content.End - 1)
    
    With MyRange
        pn = MyRange.Paragraphs.Count
        'For i = 1 To PN - 1

        T = MyRange.Paragraphs(1).Range.Text
        L = Len(T)
        ReDim Preserve PS(1)
        PS(0) = 1


        S2 = LTrim(T)
        p2 = InStr(1, S2, " ", 1)
        
        Do Until p2 = 0
        If InStr(1, S2, " ", 1) > 0 Then
        p2 = InStr(1, S2, " ", 1)
        S3 = Trim(Left(S2, p2 - 1))
        S4 = Right(S2, Len(S2) - p2)
        
        Pi = InStr(1, T, S4, 1)
        ReDim Preserve PS(UBound(PS) + 1)
        PS(UBound(PS) - 1) = Pi - 1
        
        End If
        S2 = Trim(S4)
        p2 = InStr(1, S2, " ", 1)
        
        Loop
    End With
    
    If PS(UBound(PS)) = 0 Then
        PS(UBound(PS)) = Len(T)
    End If
         'MsgBox UBound(PS)

    With MyRange
    For Each MyPara In MyRange.Paragraphs
        T = MyPara.Range.Text
        If Len(T) = 1 Then
        Exit For
        End If
        S1 = ""
        s = S1 & Chr(9)
        For j = 1 To UBound(PS)
            ST1 = Right(T, Len(T) - PS(j - 1) + 1)
            ST2 = Left(ST1, PS(j) - PS(j - 1))
            s = s & Trim(ST2) & Chr(9)
        Next j
        
        With MyRange1
        .InsertAfter Left(s, Len(s) - 1) & Chr(13)
        End With
        'MyPara.Range.Text = Left(s, Len(s) - 1) & Chr(13)
     Next MyPara
     End With
     MyRange.Select
     Selection.Delete
        Selection.WholeStory
        Selection.EndKey Unit:=wdStory
        Selection.TypeBackspace
        Selection.WholeStory
        pn = MyRange.Paragraphs.Count
        Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=UBound(PS) + 1, _
        NumRows:=pn, AutoFitBehavior:=wdAutoFitWindow

End Sub

Sub 首字母大写()
    Dim RNG As Range
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If
    Set RNG = Selection.Range
    With RNG
        .Text = LCase(RNG.Text)
        For Each aword In RNG.Words
            aword.Characters(1).Case = wdUpperCase
        Next aword
    End With
    
    '介词小写
    Dim C, T As Variant
    C = Array("Of", "In", "On", "And", "By", "At", "The", "Its", "For", "Above", "Or", "With", "Up", "To", "Under", "Through", "Over", "Into", "As", "Against")
    T = Array("of", "in", "on", "and", "by", "at", "the", "its", "for", "above", "or", "with", "up", "to", "under", "through", "over", "into", "as", "against")

    With RNG
    For i = 0 To UBound(C)
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(32) & C(i) & Chr(32)
        .Replacement.Text = Chr(32) & T(i) & Chr(32)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    End With

End Sub
Sub 介词小写()
    Dim RNG As Range
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If
    Set RNG = Selection.Range
    Dim C, T As Variant
    C = Array("Of", "In", "On", "And", "By", "At", "The", "Its", "For", "Above", "Or", "With", "Up", "To", "Under", "Through", "Over", "Into", "As", "Against")
    T = Array("of", "in", "on", "and", "by", "at", "the", "its", "for", "above", "or", "with", "up", "to", "under", "through", "over", "into", "as", "against")

    With RNG
    For i = 0 To UBound(C)
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = T(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    End With

End Sub

Sub 全部小写()
    Dim RNG As Range
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If
    Set RNG = Selection.Range
    With RNG
        .Text = LCase(RNG.Text)
    End With
        
End Sub
Sub 全部大写()
    Dim RNG As Range
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If
    Set RNG = Selection.Range
    With RNG
        .Text = UCase(RNG.Text)
    End With
        
End Sub

Sub 文本W()
    '本宏命令用途：转换文本并设定格式
    
    '把文档内容转换成为纯文本
    Selection.WholeStory
    Selection.Cut
    Selection.Collapse Direction:=wdCollapseStart
    Selection.Range.PasteSpecial DataType:=wdPasteText
    
    '将字体设等为五号宋体，行间距设定为固定16磅
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.WholeStory
        Selection.EndKey Unit:=wdStory
    Selection.WholeStory
    With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .name = "Times New Roman"
        .Size = 10.5
    End With
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 16
    End With

    '转换：软回车变为硬回车，删除空格，多余空行
    Dim O As Variant
    Dim R As Variant
    O = Array("^l", " ", "^p^p", "　", "^p", "　　^p")
    R = Array("^p", "", "^p", "", "^p　　", "")
    
    For i = 0 To 5
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = O(i)
        .Replacement.Text = R(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    
    Selection.WholeStory

'标题行加粗
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
    
    '将全角数字转换为半角数字
    Selection.WholeStory
    Dim C As Variant
    Dim D As Variant
    C = Array("０", "１", "２", "３", "４", "５", "６", "７", "８", "９", "．")
    D = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".")
    For i = 0 To 10
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = D(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    

    '每个段落之间空一行
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Paragraphs(1).Range.Select
    Selection.Delete Unit:=wdCharacter, Count:=1
    
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
        Selection.TypeBackspace
        Selection.TypeBackspace
        Selection.WholeStory
        Selection.Copy
        Selection.HomeKey Unit:=wdStory
End Sub

Sub 党务()
    '如果没有选择文本，提示用户选择
    If Len(Selection.Range.Text) = 0 Then
        ActiveDocument.Paragraphs(1).Range.Select
'        MsgBox "【注意】没有选择作为另存文档文件名的文本" & Chr(13) & _
'           "请选择或键入文件名后再选中， " & Chr(13) & _
'           "然后再执行本宏，谢谢！"
        Else
            Selection.Copy
            Set MyDOC = Application.ActiveWindow.Document
            ' 新建一个空白文档
            Documents.Add DocumentType:=wdNewBlankDocument
            Set MyDocN = Application.ActiveWindow.Document
            Application.Run MacroName:="粘贴文本"
            Application.Run MacroName:="回车替换"
            Application.Run MacroName:="删除空格K"
            Application.Run MacroName:="删除中文空格"
            NM = Left(MyDocN.Paragraphs(1).Range.Text, Len(MyDocN.Paragraphs(1).Range.Text) - 1)
            MyDOC.Activate
            ChangeFileOpenDirectory "D:\00 F2013\02 党建工作\"
            ActiveDocument.SaveAs FileName:=NM & ".doc", FileFormat:=wdFormatDocument
            MyDocN.Close SaveChanges:=wdDoNotSaveChanges
            
    End If
        ChangeFileOpenDirectory "D:\00 F2013\"
        Documents(NM & ".doc").Activate


End Sub
Sub A01_简报另存()

    ChangeFileOpenDirectory "D:\00 F2013\02 党建工作"

    NM1 = ActiveDocument.Paragraphs(1).Range.Text
    NM1 = Left(NM1, Len(NM1) - 1)
    
    NM2 = ActiveDocument.Shapes(1).TextFrame.TextRange.Text
    NM2 = Left(NM2, Len(NM2) - 1)
    
    NM = NM1 & NM2
    ActiveDocument.SaveAs FileName:=NM & ".doc", FileFormat:=wdFormatDocument
    ChangeFileOpenDirectory "D:\00 F2013\"
    
End Sub

Sub A01_简报另存1()

    ChangeFileOpenDirectory "D:\00 F2013\02 党建工作"

    NM1 = ActiveDocument.Paragraphs(3).Range.Text
    NM1 = Left(NM1, Len(NM1) - 1)
    
    NM2 = ActiveDocument.Shapes(1).TextFrame.TextRange.Text
    NM2 = Left(NM2, Len(NM2) - 1)
    
    NM = NM1 & NM2
    ActiveDocument.SaveAs FileName:=NM & ".doc", FileFormat:=wdFormatDocument
    ChangeFileOpenDirectory "D:\00 F2013\"
    
End Sub


Sub 插入翻译控制单()

    ZS = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharacters)
    ZS1 = Round(ZS / 1000, 2)
    If ZS1 < 1 Then
        qzs = "0" & ZS1
        Else
        qzs = ZS1
    End If
    
        ChangeFileOpenDirectory "D:\F2006\翻译稿件"
    
    '确定文件名
    
    Application.DisplayAlerts = wdAlertsNone
    M = IIf(Month(Date) < 10, "0" & Month(Date), Month(Date))
    D = IIf(Day(Date) < 10, "0" & Day(Date), Day(Date))
    y = Year(Date)
    N = "G" & y & M & D & "-"

Set fs = Application.FileSearch
With fs
    .LookIn = "D:\F2006\翻译稿件"
    .FileName = N
    
    If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
        FN = N & 1
    Else
        FN = N & .FoundFiles.Count + 1
    End If
End With
'MsgBox Fn

    Selection.HomeKey Unit:=wdStory
    Selection.InsertFile FileName:="C:\00 Word_dot\翻译工作流程控制单.doc", ConfirmConversions:=False
    ActiveDocument.Tables(1).Cell(1, 2).Select
    Selection.TypeText Text:=FN
    ActiveDocument.Tables(1).Cell(1, 4).Select
    T = ActiveDocument.Tables(2).Cell(1, 1).Range.Text
    nt = Left(T, Len(T) - 2)
    Selection.TypeText Text:=nt
    ActiveDocument.Tables(1).Cell(1, 6).Select
    
    Selection.TypeText Text:=qzs
    ActiveDocument.Tables(1).Cell(5, 2).Select
    Selection.TypeText Text:=Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    ActiveDocument.Tables(1).Cell(6, 2).Select
    Selection.TypeText Text:=Year(Date) & "年" & Month(Date) & "月" & Day(Date) + 1 & "日"
    
    Application.Run MacroName:="Normal.NewMacros.页面设置"

    ActiveDocument.SaveAs FileName:=FN, FileFormat:=wdFormatDocument

End Sub

Sub 所选段落格式标准化()
    
    Dim MyRange As Range
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If

    Set MyRange = Selection.Range

    With MyRange
    For Each MyPara In MyRange.Paragraphs
        T = Trim(MyPara.Range.Text)
        If Len(T) = 1 Then
        Exit For
        End If
        S1 = "　　" & T
        MyPara.Range.Text = S1
     Next MyPara
     End With

End Sub
Sub 文本格式()
    Application.Run MacroName:="删段前空"
    Application.Run MacroName:="删空行"
    Application.Run MacroName:="段前加空"
    Selection.WholeStory
    'Selection.Style = ActiveDocument.Styles("C正文")
    Application.Run MacroName:="加空行"
        Selection.WholeStory
        Selection.HomeKey Unit:=wdStory
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Application.Run MacroName:="标题粗体C"
        Selection.WholeStory
        Selection.EndKey Unit:=wdStory
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.Delete Unit:=wdCharacter, Count:=1
End Sub
Sub 表格自动调整()
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
'    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.MoveRight Unit:=wdCharacter, Count:=1

'    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
        Selection.MoveRight Unit:=wdCharacter, Count:=1

'    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
End Sub
Sub 条文加粗()

'查找特定的字符串，如："条"，然后将该字符串前的内容加粗
    T = "条　"
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    For j = 0 To ActiveDocument.Paragraphs.Count - 1
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=T
        If .Found = True Then
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            Selection.Font.Bold = True
        End If
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next j
    
        Z = "章　"
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    For j = 0 To ActiveDocument.Paragraphs.Count - 1
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=Z
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


End Sub
Sub 天津()

    '对把每一行都当成一个段落的文档标准化，删除多余的段落标记
    Selection.WholeStory
    
    '把一个段落标记后跟两个中文空格的地方更换为"※"
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p　　"
        .Replacement.Text = "※"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    '删除所有段落标记
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = ""
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
   '把"※"再转换为段落标记
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "※"
        .Replacement.Text = "^p　　"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub
Sub 段落顺序颠倒()

    Dim PS() As Variant
    pn = ActiveDocument.Paragraphs.Count
    ReDim PS(pn - 1)
    For j = 1 To pn
    PS(j - 1) = ActiveDocument.Paragraphs(j).Range.Text
    Next j
    Selection.WholeStory
    Selection.Delete
    
    For i = 0 To UBound(PS)
    Selection.TypeText Text:=PS(UBound(PS) - i)
    Next i
    
    Application.Run MacroName:="删空行"

End Sub

Sub MyFilePrint()
pass$ = InputBox("请输入打印密码：")
If pass$ = "abcd" Then
Dialogs(wdDialogFilePrint).Show
DName = ActiveDocument.Path + "\" + ActiveDocument.name
If ActiveDocument.Path = "" Then DName = "未保存文档"
Tim = str(Date) + " 日 " + str(Time)
Open "c:\print.txt" For Append As #1
Print #1, "于 " + Tim + " 打印 " + DName
Close #1
Else
MsgBox ("密码错误，请与管理人员联系！")
End If
End Sub


Sub CPFL()
Dim SourceFile, DestinationFile
SourceFile = "C:\00 Word_Dot\richeng.doc"    ' 指定源文件名。
DestinationFile = "C:\00 Word_Dot\richeng.bak"    ' 指定目的文件名。
FileCopy SourceFile, DestinationFile    ' 将源文件的内容复制到目的文件中。
End Sub

 Sub GetInput()

      ' Declare variables.
      Dim AssistantName As String
      Dim IsVisible As Boolean
      Dim Result As Byte
      Dim ball As Balloon

      ' For error trapping.
      On Error Resume Next
      Err.Clear

      ' Get the name of the current assistant.
      AssistantName = Assistant.name

      ' If the Assistant is not visible make visible.
      If Assistant.Visible = False Then
         Assistant.Visible = True
         IsVisible = False
      Else
         IsVisible = True
      End If

      ' Create a balloon for the assistant.
      Set ball = Assistant.NewBalloon

      With ball

         ' Add heading and question.
         .Heading = "Hi! I Am " & AssistantName
         .Text = "Which Animation would you like me to perform?"

         ' Add radio button choices for animation selection.
         .Labels(1).Text = "Appear"
         .Labels(2).Text = "Disappear"
         .Labels(3).Text = "Empty Trash"
         .Labels(4).Text = "Artsy"
         .Labels(5).Text = "Thinking"

         ' Sets the BalloonType Property.
         .BalloonType = msoBalloonTypeButtons

         ' Make the balloon modal, this is the default.
         .Mode = msoModeModal

         ' Add a cancel button to the balloon, OK is default.
         .Button = msoButtonSetCancel

      End With

      ' Loop until cancel is selected.
      Do

         ' Show the Balloon
         Result = ball.Show

         ' If Cancel button selected, end the macro.
         If Err <> 0 Then

            ' If the assistant was not visible close the assistant.
            If IsVisible = False Then
               Assistant.Visible = False
            Else
               ' Set to idle.
               Assistant.Animation = msoAnimationIdle
            End If

            End
         End If

         ' Perform the animation.
         Select Case Result
            Case 1
               Assistant.Animation = msoAnimationAppear
            Case 2
               Assistant.Animation = msoAnimationDisappear
            Case 3
               Assistant.Animation = msoAnimationEmptyTrash
            Case 4
               Assistant.Animation = msoAnimationGetArtsy
            Case 5
               Assistant.Animation = msoAnimationThinking
            Case Else
               MsgBox "An Error Occurred"
               End
         End Select

         ' Update the heading.
         ball.Heading = "Please Make a Selection"

      Loop

   End Sub
   Sub Main1()
      Rem Scroll line to top of Window.
      If ViewNormal() = 0 Then ViewNormal
      ScreenUpdating 0
      StartOfLine
      bmk$ = "bmk"
      i = 1
      While ExistingBookmark(bmk$) = -1
         bmk$ = "bmk" + LTrim$(str$(i))
         i = i + 1
      Wend
      EditBookmark .name = bmk$, .Add
      test = StartOfWindow()
      If test = 0 Then GoTo done
      StartOfLine
      N = 0
      While CmpBookmarks("\sel", bmk$)
         LineDown
         StartOfLine
         N = N + 1
      Wend
      VLine N
done:
      EditBookmark .name = bmk$, .Delete
   End Sub
                
Sub 翻译文稿()

    Dim NM As String
    Dim P, pn As Integer
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    Application.Run MacroName:="Normal.NewMacros.页面设置"

    pn = ActiveDocument.Content.Paragraphs.Count
    PT = 60
    If pn < PT Then
    Selection.WholeStory
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, Count:=1
    For j = 1 To PT - pn
    Selection.TypeParagraph
    Next j
    End If
    
    Dim PS(1 To 60) As Variant
    For i = 1 To pn
    PS(i) = Left(ActiveDocument.Paragraphs(i).Range.Text, Len(ActiveDocument.Paragraphs(i).Range.Text) - 1)
'    MsgBox PS(i)
    Next i
   
    Selection.WholeStory
    Selection.Cut
    Set mytable = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=60, NumColumns:=2)
    Num = 1
    For Each aCell In ActiveDocument.Tables(1).Columns(1).Cells
    aCell.Range.Text = PS(Num)
        Num = Num + 1
    Next aCell
    
    Selection.WholeStory
    'Application.Run MacroName:="Normal.NewMacros.回车替换"

    ActiveDocument.Tables(1).Columns(1).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(1).PreferredWidth = 40
    ActiveDocument.Tables(1).Columns(2).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(2).PreferredWidth = 60
    Selection.HomeKey Unit:=wdStory
    'Application.Run MacroName:="Normal.NewMacros.删除空格K"

End Sub

Sub test()
'    Dim PN As Integer
'    PN = ActiveDocument.Content.Paragraphs.Count
'    Set myTable = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=PN, NumColumns:=2)
    Num = 90
For Each aCell In ActiveDocument.Tables(1).Columns(1).Cells
    aCell.Range.Text = Num & " Sales"
    Num = Num + 1
Next aCell

End Sub



Sub Macro11()
    Selection.WholeStory
    Selection.ConvertToTable Separator:=wdSeparateByParagraphs, NumColumns:=1, _
         NumRows:=ActiveDocument.Content.Paragraphs.Count, AutoFitBehavior:=wdAutoFitFixed
    Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    ActiveDocument.Tables(1).Columns(1).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(1).PreferredWidth = 40
    ActiveDocument.Tables(1).Columns(2).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(2).PreferredWidth = 60
    Selection.HomeKey Unit:=wdStory
'    Selection.Delete
'    Selection.Delete
'    Selection.TypeParagraph
'    Application.Run macroName:="Normal.NewMacros.插入翻译控制单"
    

End Sub

Sub macro00()
    A00_网页格式
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.TypeParagraph

End Sub

Sub Macro12()
    Selection.Delete
    Selection.Delete
    ActiveDocument.Tables(1).Select
    Selection.Cut
    Selection.TypeParagraph
    Selection.Paste
    Selection.HomeKey Unit:=wdStory
    
    Application.Run MacroName:="Normal.NewMacros.插入翻译控制单"

 End Sub
 
 Sub Macro0()
 
    A00_网页格式
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.WholeStory
    
    Selection.ConvertToTable Separator:=wdSeparateByParagraphs, NumColumns:=1, _
         NumRows:=ActiveDocument.Content.Paragraphs.Count, AutoFitBehavior:=wdAutoFitFixed
    Selection.Cells.Split NumRows:=1, NumColumns:=2, MergeBeforeSplit:=False
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    ActiveDocument.Tables(1).Columns(1).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(1).PreferredWidth = 40
    ActiveDocument.Tables(1).Columns(2).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(2).PreferredWidth = 60
    Selection.HomeKey Unit:=wdStory
    
    Selection.Delete
    Selection.Delete
    ActiveDocument.Tables(1).Select
    Selection.Cut
    Selection.TypeParagraph
    Selection.Paste
    Selection.HomeKey Unit:=wdStory
    Application.Run MacroName:="Normal.NewMacros.页边距2厘米"
    
'    Application.Run MacroName:="Normal.NewMacros.插入翻译控制单"
 
 End Sub
 
Sub FNtest()
    Application.DisplayAlerts = wdAlertsNone
    N = "F" & Month(Date) & Day(Date) & "-"
    j = 1

Set fs = Application.FileSearch
With fs
    .LookIn = "D:\00 F2005\04 翻译稿件"
    .FileName = N & j
    
    If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
        FN = N & j
    Else
        FN = N & .FoundFiles.Count + 1
    End If
End With
'MsgBox Fn
 
End Sub
Sub 目录()
    Application.DisplayAlerts = wdAlertsNone

Set fs = Application.FileSearch
With fs
    .LookIn = "D:\yb2000e\ok\"
    .FileName = "*.htm"
    If .Execute(SortBy:=msoSortByFileName, _
    SortOrder:=msoSortOrderAscending) > 0 Then
        For i = 1 To .FoundFiles.Count
    Documents.Open FileName:=.FoundFiles(i)
'    ActiveDocument.Tables(1).Select
'    Selection.MoveDown unit:=wdLine, Count:=1, Extend:=wdExtend
'    Selection.Delete unit:=wdCharacter, Count:=1
'    Title = ActiveDocument.Tables(1).Cells(1).Range.Text
    ActiveDocument.Tables(1).Select
    Title = Selection.Cells(1).Range.Text
    Title = Left(Title, Len(Title) - 2)
    
    Dim NM As String
' 取得当前文件名
    Set MyDOC = Application.ActiveWindow.Document
'    MsgBox myDoc
    NM = Left(MyDOC, Len(MyDOC) - 4)
'    ActiveDocument.SaveAs FileName:=NM & "1" & ".htm", FileFormat:=wdFormatHTML

  '      Application.DisplayAlerts = wdAlertsNone
  '  ActiveDocument.Close SaveChanges:=wdSaveChanges
     ActiveWindow.Close wdDoNotSaveChanges
     
'    For Each doc In Documents
'        If InStr(1, doc.Name, "yb99e_ml.doc", 1) Then
'            doc.Activate
'            docFound = True
'            Exit For
'        Else
'            docFound = False
'        End If
'    Next doc

'    If docFound = False Then Documents.Open FileName:="D:\00 F2006\yb98c_ml.doc"

    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.InsertAfter Text:=Title & Chr(9) & NM & ".htm"

        Next i
    Else
        MsgBox "没找到可用文档"
    End If
End With
 
End Sub


Sub 超链接()

    ChangeFileOpenDirectory "C:\Users\thtfpc\Desktop"
    
    '确定文件名
    Application.DisplayAlerts = wdAlertsNone
    M = IIf(Month(Date) < 10, "0" & Month(Date), Month(Date))
    D = IIf(Day(Date) < 10, "0" & Day(Date), Day(Date))
    N = "F" & M & D & "-"

Set fs = Application.FileSearch
With fs
    .LookIn = "C:\Users\thtfpc\Desktop"
    .FileName = N
    
    If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
        FNM = N & 1
    Else
        FNM = N & .FoundFiles.Count + 1
    End If
End With


    Dim ML, FN As String
    Set mytable = ActiveDocument.Tables(1)
    R = mytable.Rows.Count
    C = mytable.Columns.Count
'    MsgBox R & " " & C
    
    ML = ""
    j = 1
    For j = 1 To R
    pn = mytable.Rows(j).Cells(2).Range.Text
    FN1 = Left(pn, Len(pn) - 2)
    mytable.Rows(j).Cells(1).Select
    td = Selection.Range.Text
    TD1 = Left(td, Len(td) - 2)
    Selection.Delete
    
    Selection.Hyperlinks.Add Anchor:=Selection.Range, Address:=ML & FN1, TextToDisplay:=TD1
   
Next j

'ChangeFileOpenDirectory "C:\Documents and Settings\ZLG005\桌面"

Selection.WholeStory
Selection.HomeKey Unit:=wdStory
ActiveDocument.Tables(1).Columns(2).Select
Selection.Cut

    ActiveDocument.SaveAs FileName:=FNM & ".htm", FileFormat:=wdFormatHTML

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Add DocumentType:=wdNewBlankDocument
ChangeFileOpenDirectory "D:\00 F2006\"

End Sub

Sub 转换表格()
    
    Dim MyRange As Range
    Dim PS() As Variant
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If

    Set MyRange = Selection.Range
    Set MyRange1 = ActiveDocument _
    .Range(start:=ActiveDocument.Content.End - 1)
    
    With MyRange
        pn = MyRange.Paragraphs.Count
        'For i = 1 To PN - 1

        T = MyRange.Paragraphs(1).Range.Text
        L = Len(T)
        ReDim Preserve PS(1)
        PS(0) = 1


        txt1 = "|"
        S2 = LTrim(T)
        p2 = InStr(1, S2, txt1, 1)
        
        Do Until p2 = 0
        If InStr(1, S2, txt1, 1) > 0 Then
        p2 = InStr(1, S2, txt1, 1)
        S3 = Trim(Left(S2, p2 - 1))
        MsgBox S3
        S4 = Right(S2, Len(S2) - p2)
        MsgBox S4

        
        Pi = InStr(1, T, S4, 1)
        ReDim Preserve PS(UBound(PS) + 1)
        PS(UBound(PS) - 1) = Pi - 1
        
        End If
        S2 = Trim(S4)
        p2 = InStr(1, S2, txt1, 1)
        
        Loop
    End With
    
    If PS(UBound(PS)) = 0 Then
        PS(UBound(PS)) = Len(T)
    End If
         'MsgBox UBound(PS)

    With MyRange
    For Each MyPara In MyRange.Paragraphs
        T = MyPara.Range.Text
        If Len(T) = 1 Then
        Exit For
        End If
        S1 = ""
        s = S1 & Chr(9)
        For j = 1 To UBound(PS)
            ST1 = Right(T, Len(T) - PS(j - 1) + 1)
            ST2 = Left(ST1, PS(j) - PS(j - 1))
            s = s & Trim(ST2) & Chr(9)
        Next j
        
        With MyRange1
        .InsertAfter Left(s, Len(s) - 1) & Chr(13)
        End With
        'MyPara.Range.Text = Left(s, Len(s) - 1) & Chr(13)
     Next MyPara
     End With
     MyRange.Select
     Selection.Delete
        Selection.WholeStory
        Selection.EndKey Unit:=wdStory
        Selection.TypeBackspace
        Selection.WholeStory
        pn = MyRange.Paragraphs.Count
        Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=UBound(PS) + 1, _
        NumRows:=pn, AutoFitBehavior:=wdAutoFitWindow

End Sub
Sub 竖线转换为制表位()

'*****************************************************
'本宏命令的功能：将竖线替换为制表位，并删除多余的空格
'作者：戴宏国
'日期：2005年9月7日
'*****************************************************

    '定义一个范围（Range）变量
    Dim MyRange As Range
    '如果没有选择范围，则指定范围为整个文档
    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    '设定范围变量为选择的范围
    Set MyRange = Selection.Range
    '取得所选区域的段落总数
    pn = MyRange.Paragraphs.Count

        '设定分隔字符
        txt1 = "|"
        '对选择的范围进行相关操作
        With MyRange
            '设定一个条件循环，j为计数器，从选定范围的第一个段落开始，循环到最后一个段落
            For j = 1 To pn
                '定义变量t，代表段落中的文本(不包括回车）
                t1 = MyRange.Paragraphs(j).Range.Text
                T = Left(t1, Len(t1) - 1)
                '定义变量p2为"|"第一次在段落中出现的位置，如果段落中没有，则p2=0
                p2 = InStr(1, T, txt1, 1)
                '定义变量s2，代表文本的总长度（不计算前导空格）
                S2 = LTrim(T)
                
                '定义一个循环，如果文本中有"|"，则找到其位置，并将其用制表位Chr(9)替换，直到全部替换为止
                Do Until p2 = 0
                
                    '判断文本中是否有"|"
                    If InStr(1, S2, txt1, 1) > 0 Then
                        p2 = InStr(1, S2, txt1, 1)
                        '定义S3为竖线前的文本，S4为竖线后的文本
                        S3 = Trim(Left(S2, p2 - 1))
                        S4 = Right(S2, Len(S2) - p2)
                    End If
                    S2 = Trim(S3) & Chr(9) & Trim(S4)
                    p2 = InStr(1, S2, txt1, 1)
                Loop
            '替换段落中的文本
            MyRange.Paragraphs(j).Range.Text = S2 & Chr(13)
        '跳转到下一个段落
        Next j
    End With
End Sub

Sub 转换为黑猫文件()

If Selection.Information(wdWithInTable) = True Then
    Selection.Copy
  Else
    MsgBox "【注意】插入点不在表格中！" & Chr(13) & _
           "　　　　请将插入点放到表格的任意单元" & Chr(13) & _
           "　　　　格中， 然后再执行本宏，谢谢！"
   End If

' 新建一个空白文档
    Documents.Add DocumentType:=wdNewBlankDocument
    Set New_Doc = Application.ActiveWindow.Document
    Selection.Paste
    Set mytable = ActiveDocument.Tables(1)
    R = mytable.Rows.Count
    C = mytable.Columns.Count
    RT = "PR('"
    For i = 1 To R
    For j = 1 To C
    pn = mytable.Rows(i).Cells(j).Range.Text
    If Len(pn) - 2 = 0 Then
    ct = "　"
    Else
    ct = Left(pn, Len(pn) - 2)
    End If

    If j < C Then
    RT = RT & ct & "','"
    Else
    RT = RT + ct + "');"
    End If
    Next j
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.InsertAfter Text:=RT
    
    RT = "PR('"
    
    Next i
    ActiveDocument.Tables(1).Select
    Selection.Cut
    Selection.Delete
    
End Sub
Sub JS打印行函数()

If Selection.Information(wdWithInTable) = True Then
    Selection.Copy
  Else
    MsgBox "【注意】插入点不在表格中！" & Chr(13) & _
           "　　　　请选择表格的任意一行" & Chr(13) & _
           "　　　　然后再执行本宏，谢谢！"
   End If
    C = Selection.Tables(1).Columns.Count
    B1 = Chr(34) & "<TR><TD width = '" & 100 / C & "%' align = 'Center'>" & Chr(34) & ","
    B2 = "," & Chr(34) & "</TD><TD width = '" & 100 / C & "%'align = 'Right'>" & Chr(34) & ","
    B3 = "," & Chr(34) & "</TD></TR>\n" & Chr(34) & ");"
    f1 = "function PR("
    f2 = "document.write("
    RT = f2 & B1
    
    For j = 1 To C
    If j < C Then
    f1 = f1 & "L" & j & ","
    Else
    f1 = f1 & "L" & j & "){"
    End If
    Next j
    
    For j = 1 To C
    If j < C Then
    RT = RT & "L" & j & B2
    Else
    RT = RT & "L" & j & B3
    End If
    Next j
    
    ' 新建一个空白文档
    Set Old_Doc = Application.ActiveWindow.Document
    Documents.Add DocumentType:=wdNewBlankDocument
    Set New_Doc = Application.ActiveWindow.Document

    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.TypeText Text:="<HTML>"
    Selection.TypeParagraph
    Selection.TypeText Text:="<HEAD>"
    Selection.TypeParagraph
    Selection.TypeText Text:="<TITLE>Type Title Here!</TITLE>"
    Selection.TypeParagraph
    Selection.TypeText Text:="<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
    Selection.TypeParagraph
    Selection.TypeText Text:="<SCRIPT language = 'JAVASCRIPT'>"
    Selection.TypeParagraph
    
    Selection.TypeText Text:=f1
    Selection.TypeParagraph
    Selection.TypeText Text:=RT
    Selection.TypeParagraph
    
    Selection.TypeText Text:="}"
    Selection.TypeParagraph
    Selection.TypeText Text:="</SCRIPT>"
    Selection.TypeParagraph
    Selection.TypeText Text:="</HEAD>"
    Selection.TypeParagraph
    Selection.TypeText Text:="<BODY>"
    Selection.TypeParagraph
    Selection.TypeText Text:="<TABLE>"
    Selection.TypeParagraph
    Selection.TypeText Text:="<SCRIPT language = 'JAVASCRIPT'>"
    Selection.TypeParagraph
    
    
    
    Selection.TypeParagraph
    Selection.TypeText Text:="</SCRIPT></TABLE></BODY></HTML>"

End Sub

Sub 删括号中的内容()
    
'------------------------------------------------------------------
'本宏命令的功能：删除括号内的内容，并删除多余的空格及段前段后空格
'作者：戴宏国
'日期：2005年11月25日
'------------------------------------------------------------------

    '定义一个范围（Range）变量
    Dim MyRange As Range
    '如果没有选择范围，则指定范围为整个文档
    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    '设定范围变量为选择的范围
    Set MyRange = Selection.Range
    '取得所选区域的段落总数
    pn = MyRange.Paragraphs.Count

        '设定分隔字符
'        txt1 = Chr(40) '(
'        txt2 = Chr(41) ')
        txt1 = Chr(60) '<
        txt2 = Chr(62) '>
'        txt1 = Chr(123) '{
'        txt2 = Chr(125) '}
        
        '对选择的范围进行相关操作
        With MyRange
            '设定一个条件循环，j为计数器，从选定范围的第一个段落开始，循环到最后一个段落
            For j = 1 To pn
                '定义变量t，代表段落中的文本(不包括回车）
                t1 = MyRange.Paragraphs(j).Range.Text
                T = Left(t1, Len(t1) - 1)
                '定义两个变量：p1,p2,分别代表"("和")"第一次在段落中出现的位置，如果段落中没有，则为0
                P1 = InStr(1, T, txt1, 1)
                p2 = InStr(1, T, txt2, 1)
                '定义变量s2，代表文本的总长度（不计算前导空格）
                S2 = LTrim(T)
                
                '定义一个循环，如果文本中有("和")"，则找到其位置，并删之，直到全部删除为止
                Do Until P1 = 0 Or p2 = 0
                
                    '判断文本中是否有("和")"
                    If InStr(1, S2, txt1, 1) > 0 And InStr(1, S2, txt2, 1) > 0 Then
                        P1 = InStr(1, S2, txt1, 1)
                        p2 = InStr(1, S2, txt2, 1)
                        If p2 < P1 Then
                        p2 = InStr(P1, S2, txt2, 1)
                       End If
                        '定义S3为("前的文本，S4为")"后的文本
                        S3 = Trim(Left(S2, P1 - 1))
                        S4 = Right(S2, Len(S2) - p2)
                    End If
                    S2 = Trim(S3) & Trim(S4)
                    P1 = InStr(1, S2, txt1, 1)
                    p2 = InStr(1, S2, txt2, 1)
                Loop
            '替换段落中的文本
            MyRange.Paragraphs(j).Range.Text = S2 & Chr(13)
        '跳转到下一个段落
        Next j
    End With
    Application.Run MacroName:="Normal.NewMacros.EDC"
    Application.Run MacroName:="Normal.NewMacros.删空行"

End Sub
Sub Folder()
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder("C:\")
    Set fc = f.Files
    For Each f1 In fc
        s = s & f1.name
        s = s & vbCrLf
    Next
    MsgBox s
End Sub

Sub 查ASCII码()
    If Len(Selection.Range.Text) = 0 Then
        MsgBox "【注意】请先选定字符，然后再执行本宏命令，谢谢！"
    Else
        MsgBox Asc(Selection.Range.Text)
    End If
End Sub


Sub Temper()
    temp = Application.InputBox(Prompt:="Please enter the temperature in degrees F.", Type:=1)
    MsgBox "The temperature is " & Celsius(temp) & " degrees C."
End Sub

Function Celsius(fDegrees)
    Celsius = (fDegrees - 32) * 5 / 9
End Function
Sub 批量删除一系列字符()
    Selection.WholeStory
    Dim A, B As Variant
    A = Array(Chr(21), Chr(22), Chr(23))
    For i = 0 To UBound(A)
        删除某个字符 (A(i))
    Selection.WholeStory
Next i
End Sub
Function 删除某个字符(txt1)
    '定义一个范围（Range）变量
    Dim MyRange As Range
    Selection.WholeStory
    If Len(Selection.Range.Text) = 0 Then
        MsgBox "【注意】请先选择，然后再执行本宏命令，谢谢！"
    Else
    Set MyRange = Selection.Range
'        txt1 = Chr(41)  '定义要删除的字符
        T = MyRange.Text
        p2 = InStr(1, T, txt1, 1)
        S2 = T
       Do Until p2 = 0
        If InStr(1, S2, txt1, 1) > 0 Then
           p2 = InStr(1, S2, txt1, 1)
           S3 = Trim(Left(S2, p2 - 1))  '定义S3为符号前的文本
           S4 = Right(S2, Len(S2) - p2) 'S4为符号后的文本
         End If
           S2 = S3 & S4
           p2 = InStr(1, S2, txt1, 1)
        Loop
            MyRange.Text = S2
    End If

End Function
Function 替换某个字符(TC)
    txt1 = Left(TC, InStr(1, TC, ",", 1) - 1)
    txt2 = Right(TC, Len(TC) - InStr(1, TC, ",", 1))
    If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
    End If
        T = Selection.Range.Text
        p2 = InStr(1, T, txt1, 1)
        S2 = T
       Do Until p2 = 0
        If InStr(1, S2, txt1, 1) > 0 Then
           p2 = InStr(1, S2, txt1, 1)
           S3 = Trim(Left(S2, p2 - 1))  '定义S3为符号前的文本
           S4 = Right(S2, Len(S2) - p2 - 1) 'S4为符号后的文本
         End If
           S2 = S3 & txt2 & S4
           p2 = InStr(1, S2, txt1, 1)
        Loop
           Selection.Range.Text = S2
End Function


Sub temp()
    Application.DisplayAlerts = wdAlertsNone
    D = "D:\txttest\"

Set fs = Application.FileSearch
With fs
    .LookIn = D
    .FileName = "*.txt"
    If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) > 0 Then
        For i = 1 To .FoundFiles.Count
        MsgBox .FoundFiles(i).name
        ChangeFileOpenDirectory D
        Selection.InsertFile FileName:=.FoundFiles(i), Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False
        Application.Run MacroName:="Normal.NewMacros.删括号中的内容"
        ActiveDocument.SaveAs FileName:=.FoundFiles(i), FileFormat:=wdFormatText, AddToRecentFiles:=False
        Document.Add
        Next i
    Else
        MsgBox "没找到可用文档"
    End If
End With
 
End Sub
Sub 批量删除括号中的内容()
    D = "D:\txttest\"
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(D)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.name
        s = Left(s, Len(s) - 4)
        ChangeFileOpenDirectory D
    '打开文件
        Documents.Open FileName:=s & ".txt", ConfirmConversions:=False, _
        ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        WritePasswordTemplate:="", Format:=wdOpenFormatAuto, Encoding:=936
    '运行宏命令
'       Application.Run MacroName:="Normal.NewMacros.A测试"
        Application.Run MacroName:="Normal.NewMacros.删括号中的内容"
        Application.Run MacroName:="Normal.NewMacros.删段前空"
        Application.Run MacroName:="Normal.NewMacros.缩空"
        Application.Run MacroName:="Normal.NewMacros.删空行"
    '保存文件，文件名后加了“n”
        ActiveDocument.SaveAs FileName:="N_" & s, FileFormat:=wdFormatText, _
        LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
        ActiveWindow.Close
    Next
End Sub
Function dc(C) '删除文档中所有指定的字符串 函数
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C
        .Replacement.Text = ""
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory

End Function
Sub EDC()
    dc ("&nbsp;")
    dc ("#")
    dc ("＃")
   ' DC (Chr(13))
End Sub



Sub jf()
'*****************************************************
'本宏命令的功能：将文档的选定内容转换为繁体
'作者：戴宏国
'日期：2005年11月25日
'*****************************************************

ML = "D:\00 Word_Dot\"
ChangeFileOpenDirectory ML


Documents.Open ("jtchar.doc")
jt = ActiveDocument.Content.Paragraphs(1).Range.Text
ActiveDocument.Close

Documents.Open ("ftchar.doc")
ft = ActiveDocument.Content.Paragraphs(1).Range.Text
ActiveDocument.Close

    '定义一个范围（Range）变量
    Dim MyRange As Range
    '如果没有选择范围，则指定范围为整个文档
    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    '设定范围变量为选择的范围
    Set MyRange = Selection.Range

With MyRange
T = MyRange.Text
s = ""
L = Len(T)
For i = 1 To L
C = MyRange.Characters(i)

N = InStr(1, jt, C, vbTextCompare)
If N > 0 Then
s = s & Right(Left(ft, N), 1)
Else
s = s & C
End If

Next i
MyRange.Text = s
End With

End Sub
Sub fj()
'*****************************************************
'本宏命令的功能：将文档的选定内容转换为简体
'作者：戴宏国
'日期：2005年11月25日
'*****************************************************

ML = "D:\00 Word_Dot\"
ChangeFileOpenDirectory ML


Documents.Open ("jtchar.doc")
jt = ActiveDocument.Content.Paragraphs(1).Range.Text
ActiveDocument.Close

Documents.Open ("ftchar.doc")
ft = ActiveDocument.Content.Paragraphs(1).Range.Text
ActiveDocument.Close

    '定义一个范围（Range）变量
    Dim MyRange As Range
    '如果没有选择范围，则指定范围为整个文档
    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    '设定范围变量为选择的范围
    Set MyRange = Selection.Range

With MyRange
T = MyRange.Text
s = ""
L = Len(T)
For i = 1 To L
C = MyRange.Characters(i)

N = InStr(1, ft, C, vbTextCompare)
If N > 0 Then
s = s & Right(Left(jt, N), 1)
Else
s = s & C
End If

Next i
MyRange.Text = s
End With

End Sub

'SubStr()    中文化取子字串，相对Mid()
'Strlen()    中文化字串长度，相对Len()
'StrLeft()   中文化取左字串，相对Left()
'StrRight()  中文化取右字串，相对Right()
'isChinese() Check某个字是否中文字

Public Function SubStr(ByVal tstr As String, start As Integer, Optional leng As Variant) As String
Dim tmpStr  As String
If IsMissing(leng) Then
   tmpStr = StrConv(MidB(StrConv(tstr, vbFromUnicode), start), vbUnicode)
Else
   tmpStr = StrConv(MidB(StrConv(tstr, vbFromUnicode), start, leng), vbUnicode)
End If
SubStr = tmpStr
End Function

Public Function Strlen(ByVal tstr As String) As Integer
   Strlen = LenB(StrConv(tstr, vbFromUnicode))
End Function

Public Function StrLeft(ByVal str5 As String, ByVal len5 As Long) As String
Dim tmpStr As String
tmpStr = StrConv(str5, vbFromUnicode)
tmpStr = LeftB(tmpStr, len5)
StrLeft = StrConv(tmpStr, vbUnicode)
End Function

Public Function StrRight(ByVal str5 As String, ByVal len5 As Long) As String
Dim tmpStr As String
tmpStr = StrConv(str5, vbFromUnicode)
tmpStr = RightB(tmpStr, len5)
StrLeft = StrConv(tmpStr, vbUnicode)
End Function

Public Function isChinese(ByVal asciiv As Integer) As Boolean
   If Len(Hex$(asciiv)) > 2 Then
      isChinese = True
   Else
      isChinese = False
   End If
End Function

Function shc(txt1)    '定义一个范围（Range）变量
    Dim MyRange As Range
    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    
    Set MyRange = Selection.Range
'        txt1 = Chr(41)  '定义要删除的字符
        T = MyRange.Text
        p2 = InStr(1, T, txt1, 1)
        S2 = T
       Do Until p2 = 0
        If InStr(1, S2, txt1, 1) > 0 Then
           p2 = InStr(1, S2, txt1, 1)
           S3 = Trim(Left(S2, p2 - 1))  '定义S3为符号前的文本
           S4 = Right(S2, Len(S2) - p2) 'S4为符号后的文本
         End If
           S2 = S3 & S4
           p2 = InStr(1, S2, txt1, 1)
        Loop
            MyRange.Text = S2

End Function
Sub 转换NJ98()
    '定义一个范围（Range）变量
    Dim MyRange As Range
    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    Set MyRange = Selection.Range
        txt1 = Chr(124)  '定义要删除的字符
        txt2 = Chr(9)
        T = MyRange.Text
        p2 = InStr(1, T, txt1, 1)
        S2 = T
       Do Until p2 = 0
        If InStr(1, S2, txt1, 1) > 0 Then
           p2 = InStr(1, S2, txt1, 1)
           S3 = Trim(Left(S2, p2 - 1))  '定义S3为符号前的文本
           S4 = Trim(Right(S2, Len(S2) - p2)) 'S4为符号后的文本
         End If
           S2 = S3 & txt2 & S4
           p2 = InStr(1, S2, txt1, 1)
        Loop
            MyRange.Text = S2
        
        shc ("--")
        shc ("-" & Chr(13))

        
        With MyRange
        pn = MyRange.Paragraphs.Count
        For i = 1 To pn - 1
        T = MyRange.Paragraphs(i).Range.Text
        T = Left(T, Len(T) - 1)
        If Left(T, 1) = Chr(9) Then
        s = Right(T, Len(T) - 1)
        Else
        s = T
        End If
        
        If Right(s, 1) = Chr(9) Then
        s = Left(s, Len(s) - 1)
        Else
        s = T
        End If
        
        MyRange.Paragraphs(i).Range.Text = s & Chr(13)
        Next i
        End With
        Application.Run MacroName:="Normal.NewMacros.缩空"
        Application.Run MacroName:="Normal.NewMacros.删空行"
        
        Selection.WholeStory
        A2 = "北 京 Beijing"
        A3 = "天 津 Tianjin"
        TXT = Selection.Range.Text
        P1 = InStr(1, TXT, A2, vbTextCompare)
        p2 = InStr(1, TXT, A3, vbTextCompare)
        If P1 > 0 And p2 > 0 Then
        Application.Run MacroName:="Normal.NewMacros.省份替换"
        End If
        
        shc ("ed by Region")


End Sub

Sub 省份替换()
Dim A, B As Variant
A1 = "全 国 National Total"
A2 = "北 京 Beijing"
A3 = "天 津 Tianjin"
A4 = "河 北 Hebei"
A5 = "山 西 Shanxi"
A6 = "内蒙古 Inner Mongolia"
A7 = "辽 宁 Liaoning"
A8 = "吉 林 Jilin"
A9 = "黑龙江 Heilongjiang"
A10 = "上 海 Shanghai"
A11 = "江 苏 Jiangsu"
A12 = "浙 江 Zhejiang"
A13 = "安 徽 Anhui"
A14 = "福 建 Fujian"
A15 = "江 西 Jiangxi"
A16 = "山 东 Shandong"
A17 = "河 南 Henan"
A18 = "湖 北 Hubei"
A19 = "湖 南 Hunan"
A20 = "广 东 Guangdong"
A21 = "广 西 Guangxi"
A22 = "海 南 Hainan"
A23 = "重 庆 Chongqing"
A24 = "四 川 Sichuan"
A25 = "贵 州 Guizhou"
A26 = "云 南 Yunnan"
A27 = "西 藏 Tibet"
A28 = "陕 西 Shaanxi"
A29 = "甘 肃 Gansu"
A30 = "青 海 Qinghai"
A31 = "宁 夏 Ningxia"
A32 = "新 疆 Xinjiang"
A33 = "不分地区 Not Classifi-"

B1 = "全国合计" & Chr(9) & "National Total"
B2 = "北　京" & Chr(9) & "Beijing"
B3 = "天　津" & Chr(9) & "Tianjin"
B4 = "河　北" & Chr(9) & "Hebei"
B5 = "山　西" & Chr(9) & "Shanxi"
B6 = "内蒙古" & Chr(9) & "Inner Mongolia"
B7 = "辽　宁" & Chr(9) & "Liaoning"
B8 = "吉　林" & Chr(9) & "Jilin"
B9 = "黑龙江" & Chr(9) & "Heilongjiang"
B10 = "上　海" & Chr(9) & "Shanghai"
B11 = "江　苏" & Chr(9) & "Jiangsu"
B12 = "浙　江" & Chr(9) & "Zhejiang"
B13 = "安　徽" & Chr(9) & "Anhui"
B14 = "福　建" & Chr(9) & "Fujian"
B15 = "江　西" & Chr(9) & "Jiangxi"
B16 = "山　东" & Chr(9) & "Shandong"
B17 = "河　南" & Chr(9) & "Henan"
B18 = "湖　北" & Chr(9) & "Hubei"
B19 = "湖　南" & Chr(9) & "Hunan"
B20 = "广　东" & Chr(9) & "Guangdong"
B21 = "广　西" & Chr(9) & "Guangxi"
B22 = "海　南" & Chr(9) & "Hainan"
B23 = "重　庆" & Chr(9) & "Chongqing"
B24 = "四　川" & Chr(9) & "Sichuan"
B25 = "贵　州" & Chr(9) & "Guizhou"
B26 = "云　南" & Chr(9) & "Yunnan"
B27 = "西　藏" & Chr(9) & "Tibet"
B28 = "陕　西" & Chr(9) & "Shaanxi"
B29 = "甘　肃" & Chr(9) & "Gansu"
B30 = "青　海" & Chr(9) & "Qinghai"
B31 = "宁　夏" & Chr(9) & "Ningxia"
B32 = "新　疆" & Chr(9) & "Xinjiang"
B33 = "不分地区" & Chr(9) & "Not Classified by Region"
A = Array(A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, A12, A13, A14, A15, A16, A17, A18, A19, A20, A21, A22, A23, A24, A25, A26, A27, A28, A29, A30, A31, A32, A33)
B = Array(B1, B2, B3, B4, B5, B6, B7, B8, B9, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19, B20, B21, B22, B23, B24, B25, B26, B27, B28, B29, B30, B31, B32, B33)

    Selection.WholeStory
    For i = 0 To 32
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = A(i)
        .Replacement.Text = B(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i

End Sub

Sub 批量转换NJ98()
    Application.DisplayAlerts = wdAlertsNone
    D = "D:\NJ98\"
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(D)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.name
        s = Left(s, Len(s) - 4)
        ChangeFileOpenDirectory D
    '打开文件
        Documents.Open FileName:=s & ".txt", ConfirmConversions:=False, _
        ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        WritePasswordTemplate:="", Format:=wdOpenFormatAuto, Encoding:=936
    '运行宏命令
'       Application.Run MacroName:="Normal.NewMacros.A测试"
        Application.Run MacroName:="Normal.NewMacros.转换NJ98"
    '保存文件，文件名后加了“C”
        ActiveDocument.SaveAs FileName:=s & "C", FileFormat:=wdFormatDocument, _
        LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
    '保存文件，文件名后加了“E”
        ActiveDocument.SaveAs FileName:=s & "E", FileFormat:=wdFormatDocument, _
        LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
        
        Selection.WholeStory
        Application.Run MacroName:="Normal.NewMacros.删汉字"
        ActiveDocument.Save
        ActiveWindow.Close
        
        Documents.Open (s & "C.doc")
        Selection.WholeStory
        Application.Run MacroName:="Normal.NewMacros.删除英文字母"
        ActiveDocument.Save
        ActiveWindow.Close
    Next f1
 
End Sub

Sub nj99()
    Application.DisplayAlerts = wdAlertsNone

Set fs = Application.FileSearch
With fs
    .LookIn = "D:\yb2000e\"
    .FileName = "*.htm"
    If .Execute(SortBy:=msoSortByFileName, _
    SortOrder:=msoSortOrderAscending) > 0 Then
        For i = 1 To .FoundFiles.Count
    Documents.Open FileName:=.FoundFiles(i)
    Selection.MoveDown Unit:=wdLine, Count:=3, Extend:=wdExtend
    Selection.Delete Unit:=wdCharacter, Count:=1
    ActiveDocument.Save
    ActiveDocument.Close

        Next i
    Else
        MsgBox "没找到可用文档"
    End If
End With
 
End Sub
Sub 删表空格()

    Set mytable = ActiveDocument.Tables(1)
    With mytable
        For Each aCell In mytable.Range.Cells
            T = aCell.Range.Text
            t1 = Left(T, 1)
            If t1 = Chr(-24159) Then
                aCell.Range.Text = ""
            End If
        Next aCell
    End With

End Sub
Sub 插入序号列()

    Set mytable = ActiveDocument.Tables(1)
    Set aColumn = mytable.Columns.Add(BeforeColumn:=mytable.Columns(1))
    For Each aCell In aColumn.Cells
        aCell.Range.Delete
        aCell.Range.InsertAfter Num + 1
        Num = Num + 1
    Next aCell
    
End Sub

Sub 删除空行空列()
If ActiveDocument.Tables.Count >= 1 Then
    Set mytable = ActiveDocument.Tables(1)
    C = ActiveDocument.Tables(1).Columns.Count
    R = ActiveDocument.Tables(1).Rows.Count
'    MsgBox R & "×" & C
End If
For i = 0 To R - 1
    H = R - i
For Each aCell In mytable.Rows(H).Cells
    X = Len(aCell.Range.Text)
 '   MsgBox X
    If X > 2 Then
    Exit For
    End If
Next aCell
If X > 2 Then
Exit For
End If

    mytable.Rows(H).Select
    Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
    Selection.Delete

Next i
End Sub
Sub 按内容调表()
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
End Sub
Sub 固定列宽()
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
End Sub
Sub 按窗口调表()
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
End Sub
Sub 页面视图()
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
End Sub
Sub 网页视图()
    ActiveWindow.View.Type = wdWebView
End Sub
Sub 普通视图()
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdNormalView
    Else
        ActiveWindow.View.Type = wdNormalView
    End If
End Sub
Sub 编表()
        Application.Run MacroName:="Normal.NewMacros.删除中文空格"
        Application.Run MacroName:="Normal.NewMacros.页面视图"
        Application.Run MacroName:="Normal.NewMacros.按内容调表"
        Application.Run MacroName:="Normal.NewMacros.固定列宽"

End Sub
Sub 全角标点转换半角标点()

    '全角标点符号转换半角标点符号
    Selection.WholeStory
    Dim C As Variant
    Dim D As Variant
    C = Array("，", "。", "：", "“", "”", "《", "》", "（", "）", Chr(37), Chr(-23643), "、", "；", "‰")
    
    D = Array(", ", ". ", ":", Chr(34), Chr(34), Chr(34), Chr(34), "(", ")", " percent", " percent", ", ", "; ", " per thousand")
    For i = 0 To UBound(C)
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = D(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    
End Sub

Sub 翻译常用统计术语()
    
    Dim MyRange As Range
    Dim f, R As Variant
    
    Dim CH(), EN() As Variant
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If
    Set MyRange = Selection.Range
    
    Documents.Open FileName:="D:\00 Word_Dot\常用统计术语英汉对照表.doc"
    H = ActiveDocument.Tables(1).Rows.Count
    For i = 1 To H
    C = ActiveDocument.Tables(1).Columns(1).Cells(i).Range.Text
    C = Left(C, Len(C) - 2)
    E = ActiveDocument.Tables(1).Columns(2).Cells(i).Range.Text
    E = Left(E, Len(E) - 2)
    ReDim Preserve CH(i)
    CH(i) = C
    ReDim Preserve EN(i)
    EN(i) = E
    
    Next i
    ActiveDocument.Close
    
    For i = 1 To H
    f = CH(i)
    R = EN(i)
    s = MyRange.Text
    P = InStr(1, s, f, 1)
    If P > 0 Then
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                 .Text = f
                 .Replacement.Text = Chr(32) & R & Chr(32)
             End With
             Selection.Find.Execute Replace:=wdReplaceAll
    End If
    Next i
    Application.Run MacroName:="Normal.NewMacros.缩空"
    Application.Run MacroName:="Normal.NewMacros.删段后空"
    Application.Run MacroName:="Normal.NewMacros.删段后空"
    Application.Run MacroName:="Normal.NewMacros.删空行"
    Application.Run MacroName:="Normal.NewMacros.翻译_亿元"
    Application.Run MacroName:="Normal.NewMacros.翻译_万吨"
    Application.Run MacroName:="Normal.NewMacros.翻译_万平方米"
    Application.Run MacroName:="Normal.NewMacros.全角标点转换半角标点"
    Application.Run MacroName:="Normal.NewMacros.缩空"
    Application.Run MacroName:="Normal.NewMacros.删空行"

    
End Sub

Sub 翻译_亿元()

    N = ActiveDocument.Paragraphs.Count
    For i = 1 To N
        T = ActiveDocument.Paragraphs(i).Range.Text
        T = Left(T, Len(T) - 1)
        t1 = "亿元"
        S2 = LTrim(T)
        p2 = InStr(1, S2, t1, 1)
        Do Until p2 = 0
        If InStr(1, S2, t1, 1) > 0 Then
        p2 = InStr(1, S2, t1, 1)
        S3 = Trim(Left(S2, p2 - 1))
        S4 = Right(S2, Len(S2) - p2 - 1)
        For j = 1 To Len(S3)
        C = Left(Right(S3, j), 1)
        If Asc(C) > 45 And Asc(C) < 58 Then
        j = j + 1
        Else
        X = j
        D = Right(S3, X - 1)
        W = Left(S3, Len(S3) - X + 1)
        If Asc(Left(D, 1)) < 45 Then
        D = Right(S3, X - 2)
        W = Left(S3, Len(S3) - X + 2)
        End If
        Exit For
        End If
        Next j
        S2 = W & str(Val(D) / 10) & " billion yuan" & S4
        End If
        p2 = InStr(1, S2, t1, 1)
        Loop
        ActiveDocument.Paragraphs(i).Range.Text = S2 & Chr(13)
    Next i

End Sub

Sub 翻译_万吨()

    N = ActiveDocument.Paragraphs.Count
    For i = 1 To N
        T = ActiveDocument.Paragraphs(i).Range.Text
        T = Left(T, Len(T) - 1)
        t1 = "万吨"
        S2 = LTrim(T)
        p2 = InStr(1, S2, t1, 1)
        
        Do Until p2 = 0
        If InStr(1, S2, t1, 1) > 0 Then
        p2 = InStr(1, S2, t1, 1)
        S3 = Trim(Left(S2, p2 - 1))
        S4 = Right(S2, Len(S2) - p2 - 1)
        
        For j = 1 To Len(S3)
            C = Left(Right(S3, j), 1)
            If Asc(C) > 45 And Asc(C) < 58 Then
                j = j + 1
            Else
                X = j
                D = Right(S3, X - 1)
                W = Left(S3, Len(S3) - X + 1)
                If Asc(Left(D, 1)) < 45 Then
                    D = Right(S3, X - 2)
                    W = Left(S3, Len(S3) - X + 2)
                End If
            Exit For
            End If
        Next j
        S2 = W & str(Val(D) / 100) & " million tons" & S4
        End If
        p2 = InStr(1, S2, t1, 1)
        Loop
        
        ActiveDocument.Paragraphs(i).Range.Text = S2 & Chr(13)
        Next i

End Sub

Sub 翻译_万平方米()

    N = ActiveDocument.Paragraphs.Count
    For i = 1 To N
        T = ActiveDocument.Paragraphs(i).Range.Text
        T = Left(T, Len(T) - 1)
        t1 = "万平方米"
        S2 = LTrim(T)
        p2 = InStr(1, S2, t1, 1)
        
        Do Until p2 = 0
            If InStr(1, S2, t1, 1) > 0 Then
                p2 = InStr(1, S2, t1, 1)
                S3 = Trim(Left(S2, p2 - 1))
                S4 = Right(S2, Len(S2) - p2 - 3)
                For j = 1 To Len(S3)
                    C = Left(Right(S3, j), 1)
                    If Asc(C) > 45 And Asc(C) < 58 Then
                        j = j + 1
                    Else
                        X = j
                        D = Right(S3, X - 1)
                        W = Left(S3, Len(S3) - X + 1)
                        If Asc(Left(D, 1)) < 45 Then
                            D = Right(S3, X - 2)
                            W = Left(S3, Len(S3) - X + 2)
                        End If
                        Exit For
                    End If
                    Next j
                S2 = W & str(Val(D) / 100) & " million square meters" & S4
            End If
            p2 = InStr(1, S2, t1, 1)
        Loop
        ActiveDocument.Paragraphs(i).Range.Text = S2 & Chr(13)
    Next i

End Sub

Sub 段落左右缩进()
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.1)
        .RightIndent = CentimetersToPoints(0.1)
    End With
End Sub

Sub AAtable()
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
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
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

    
   Else
    MsgBox "【注意】插入点不在表格中！" & Chr(13) & _
           "　　　　请将插入点放到表格的任意单元" & Chr(13) & _
           "　　　　格中， 然后再执行本宏，谢谢！"
   End If
End Sub
Sub 表格标题与单位()

    '判断是否选择了有关内容
    If Len(Selection.Range.Text) = 0 Then
        MsgBox "你没用选中任何内容！" & Chr(13) & "请选中“表格的标题、单位和表格，" & Chr(13) & "然后再运行本宏命令!"
    Else

    '判断是否选择了有关内容
    If Selection.Tables.Count = 0 Then
        MsgBox "你没用选中表格！" & Chr(13) & "请选中“表格的标题、单位和表格，" & Chr(13) & "然后再运行本宏命令!"
    Else
    
    '定义选中的区域，提取表格的标题和单位的文本
    Set MyRange = Selection.Range
    Title = MyRange.Paragraphs(1).Range.Text
    Title = Left(Title, Len(Title) - 1)
    Unit = MyRange.Paragraphs(2).Range.Text
    Unit = Left(Unit, Len(Unit) - 1)
    
    '插入单位行，并设置格式
    MyRange.Tables(1).Select
    Selection.InsertRowsAbove 1
    Selection.Cells.Merge
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.InsertAfter Text:=Trim(Unit)
    Selection.Rows.Height = CentimetersToPoints(0.6)
        With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    
    '插入标题行，并设置格式
    Selection.InsertRowsAbove 1
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.InsertAfter Text:=Trim(Title)
    Selection.Rows.Height = CentimetersToPoints(1#)
    
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone

    With Selection.Font
    .Bold = True
    .Size = 12
    End With
    
    '删除原来的文本
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    MyRange.Paragraphs(2).Range.Select
    Selection.Delete
    MyRange.Paragraphs(1).Range.Select
    Selection.Delete
    MyRange.Tables(1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=2
    Selection.SelectRow
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
    Selection.Borders(wdBorderTop).LineWidth = wdLineWidth150pt
    MyRange.Tables(1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    End If
    End If
End Sub
Sub 单位和日期()  '宏命令名称
    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(10), Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
    Selection.TypeText Text:=Chr(9) & "资料中心" & Chr(13)
    Selection.InsertDateTime DateTimeFormat:=Chr(9) & "EEEE年O月A日", InsertAsField:=False
End Sub
Sub 删除单元格中的0()
'指定处理的表格为文档中的第1张表
Set mytable = ActiveDocument.Tables(1)
'设定循环
For Each celTable In mytable.Range.Cells
'设定范围，不包括单元格内容的换行符
    Set rngTable = ActiveDocument.Range(start:=celTable.Range.start, End:=celTable.Range.End - 1)
    If rngTable.Text = "0" Then
        rngTable.Text = ""
    End If
    Next celTable
End Sub

Sub GBE_table()
'
' GBE_table Macro
' 宏在 2007-3-12 由 戴宏国: 录制
'
    Dim T As Integer
    T = ActiveDocument.Tables.Count ' 取得文档中表格的张数
    
     With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
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
        .GutterPos = wdGutterPosLeft
        .LayoutMode = wdLayoutModeLineGrid
    End With
   
    
    If T < 1 Then
        MsgBox "文档中没有表格，无法执行本宏命令!"
    Else
    
       If Selection.Information(wdWithInTable) = True Then
    Selection.Tables(1).Select
    Else
    
        ActiveDocument.Tables(1).Select
    End If
    
    Selection.Font.Size = 10
    Selection.Font.name = "Arial"
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.1)
        .RightIndent = CentimetersToPoints(0.1)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 12
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .WordWrap = True
    End With
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    
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
    Selection.Tables(1).Rows.Alignment = wdAlignRowCenter
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.4)
    
    Options.DefaultBorderLineWidth = wdLineWidth150pt
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
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    
    Selection.Tables(1).Cell(1, 1).Select
    Selection.SelectColumn
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    
    Selection.Tables(1).Cell(1, 1).Select
    Selection.SelectRow
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter

    End If
    
End Sub

'----------------------------------------------------------------------------
'本宏命令的功能：自动按不同情况对申请表进行编号、保存、登记
'作者：戴宏国
'日期：2008年5月5日
'----------------------------------------------------------------------------
Sub G01_登记()

    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Set MyDOC = Application.ActiveWindow.Document  '确定活动文档为工作文档
    
    If MyDOC.Tables.Count = 0 Then  '判断文档中是否有表格
        MsgBox ("文档中没有表格，不能执行本宏命令！")
        mycheck = False
    Else
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        If TN <> "国家统计局政府信息公开申请表" Then
            MsgBox ("非政府信息公开申请表，不能执行本宏命令！")
            mycheck = False
        Else
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
            lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '读取表格类型
            dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '读取：申请日期
            '判断是否已经登记
            If Len(bh) < 10 Then
                If IsDate(dt) Then
                    mycheck = True
                Else
                    MsgBox "申请日期格式有误，请先修改，再执行本命令！" & Chr(13) & "正确的格式为：2008-5-1，2008-05-01或2008年5月1日"
                    mycheck = False
                End If
            Else
                MsgBox ("此申请表已登记，不能重复登记！")
                mycheck = False
            End If
        End If
    End If

    Do While mycheck = True
    
    Application.DisplayAlerts = wdAlertsNone
    M = IIf(Month(Date) < 10, "0" & Month(Date), Month(Date))
    D = IIf(Day(Date) < 10, "0" & Day(Date), Day(Date))
    y = Year(Date)
    If lb = "个人表" Then
    N = "G" & y & M & D & "-"
    Else
    N = "D" & y & M & D & "-"
    End If

'在工作目录中查找是否有当天存入的文件，如果没有，编号从1开始，否则，在原来的编号基础上连续编号
Set fs = Application.FileSearch
With fs
    .LookIn = "D:\信息公开"
    .FileName = "申请" & N
    
    If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
        FN = N & 1
    Else
        FN = N & .FoundFiles.Count + 1
    End If
End With

    If Len(ActiveDocument.Tables(1).Cell(2, 2).Range.Text) < Len(FN) Then
    ActiveDocument.Tables(1).Cell(2, 2).Select
    Selection.TypeText Text:=FN
    ActiveDocument.SaveAs FileName:="申请" & FN, FileFormat:=wdFormatDocument   '保存文件
    End If
    
    If lb = "个人表" Then
    Set MyDOC = ActiveDocument
    Set MyTab = MyDOC.Tables(1)
    t1 = Left(MyTab.Cell(5, 3).Range.Text, Len(MyTab.Cell(5, 3).Range.Text) - 2)
    t2 = Left(MyTab.Cell(6, 3).Range.Text, Len(MyTab.Cell(6, 3).Range.Text) - 2)
    t3 = Left(MyTab.Cell(7, 3).Range.Text, Len(MyTab.Cell(7, 3).Range.Text) - 2)
    t4 = Left(MyTab.Cell(8, 3).Range.Text, Len(MyTab.Cell(8, 3).Range.Text) - 2)
    t5 = Left(MyTab.Cell(9, 3).Range.Text, Len(MyTab.Cell(9, 3).Range.Text) - 2)
    t6 = Left(MyTab.Cell(10, 3).Range.Text, Len(MyTab.Cell(10, 3).Range.Text) - 2)
    t7 = Left(MyTab.Cell(11, 3).Range.Text, Len(MyTab.Cell(11, 3).Range.Text) - 2)
    t8 = Left(MyTab.Cell(12, 3).Range.Text, Len(MyTab.Cell(12, 3).Range.Text) - 2)
    t9 = Left(MyTab.Cell(13, 3).Range.Text, Len(MyTab.Cell(13, 3).Range.Text) - 2)
    t10 = Left(MyTab.Cell(14, 3).Range.Text, Len(MyTab.Cell(14, 3).Range.Text) - 2)
    t11 = Left(MyTab.Cell(16, 3).Range.Text, Len(MyTab.Cell(16, 3).Range.Text) - 2)
    t12 = Left(MyTab.Cell(17, 3).Range.Text, Len(MyTab.Cell(17, 3).Range.Text) - 2)
    t13 = Left(MyTab.Cell(18, 3).Range.Text, Len(MyTab.Cell(18, 3).Range.Text) - 2)
    t14 = Left(MyTab.Cell(19, 3).Range.Text, Len(MyTab.Cell(19, 3).Range.Text) - 2)
    t15 = Left(MyTab.Cell(20, 3).Range.Text, Len(MyTab.Cell(20, 3).Range.Text) - 2)

    Documents.Open FileName:="D:\信息公开\登记表1.doc"
    Set MyDOC = ActiveDocument
    Set MyTab = MyDOC.Tables(1)
    N = MyTab.Rows.Count
    i = 0
    MyTab.Cell(N, 1).Select
    Selection.InsertRowsBelow 1

    A = Array(N, FN, "", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15)
    
    For Each aCell In MyTab.Rows(N + 1).Range.Cells
        aCell.Range.Text = A(i)
        i = i + 1
    Next aCell
    ActiveDocument.Close SaveChanges:=wdSaveChanges
'    MsgBox ("已成功登记！")
    End If
    
    If lb = "单位表" Then
    Set MyDOC = ActiveDocument
    Set MyTab = MyDOC.Tables(1)
    t1 = Left(MyTab.Cell(5, 3).Range.Text, Len(MyTab.Cell(5, 3).Range.Text) - 2)
    t2 = Left(MyTab.Cell(6, 3).Range.Text, Len(MyTab.Cell(6, 3).Range.Text) - 2)
    t3 = Left(MyTab.Cell(7, 3).Range.Text, Len(MyTab.Cell(7, 3).Range.Text) - 2)
    t4 = Left(MyTab.Cell(8, 3).Range.Text, Len(MyTab.Cell(8, 3).Range.Text) - 2)
    t5 = Left(MyTab.Cell(9, 3).Range.Text, Len(MyTab.Cell(9, 3).Range.Text) - 2)
    t6 = Left(MyTab.Cell(10, 3).Range.Text, Len(MyTab.Cell(10, 3).Range.Text) - 2)
    t7 = Left(MyTab.Cell(11, 3).Range.Text, Len(MyTab.Cell(11, 3).Range.Text) - 2)
    t8 = Left(MyTab.Cell(12, 3).Range.Text, Len(MyTab.Cell(12, 3).Range.Text) - 2)
    t9 = Left(MyTab.Cell(13, 3).Range.Text, Len(MyTab.Cell(13, 3).Range.Text) - 2)
    t10 = Left(MyTab.Cell(14, 3).Range.Text, Len(MyTab.Cell(14, 3).Range.Text) - 2)
    t11 = Left(MyTab.Cell(16, 3).Range.Text, Len(MyTab.Cell(16, 3).Range.Text) - 2)
    t12 = Left(MyTab.Cell(17, 3).Range.Text, Len(MyTab.Cell(17, 3).Range.Text) - 2)
    t13 = Left(MyTab.Cell(18, 3).Range.Text, Len(MyTab.Cell(18, 3).Range.Text) - 2)
    t14 = Left(MyTab.Cell(19, 3).Range.Text, Len(MyTab.Cell(19, 3).Range.Text) - 2)

    Documents.Open FileName:="D:\信息公开\登记表2.doc"
    Set MyDOC = ActiveDocument
    Set MyTab = MyDOC.Tables(1)
    N = MyTab.Rows.Count
    i = 0
    MyTab.Cell(N, 1).Select
    Selection.InsertRowsBelow 1

    A = Array(N, FN, "", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14)
    
    For Each aCell In MyTab.Rows(N + 1).Range.Cells
        aCell.Range.Text = A(i)
        i = i + 1
    Next aCell
    ActiveDocument.Close SaveChanges:=wdSaveChanges
'    MsgBox ("已成功登记！")
    End If
    
    Exit Do
    Loop
    
End Sub

'--------------------------------------------------------------
'本宏命令的功能：读取申请表信息、自动生成回执并保存到指定目录
'作者：戴宏国
'日期：2008年5月6日
'--------------------------------------------------------------

Sub G02_回执()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Set MyDOC = Application.ActiveWindow.Document  '确定活动文档为工作文档
    If MyDOC.Tables.Count > 0 Then  '判断文档中是否有表格
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        If TN = "国家统计局政府信息公开申请表" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
            '判断是否已经登记
            If Len(bh) < 10 Then
                MsgBox ("申请表尚未登记，请先登记！")
                mycheck = False
            Else
                '判断是否已经生成过登记回执
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\信息公开"
                        .FileName = "回执" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("已经生成过登记回执，不能再次生成！")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("非政府信息公开申请表，不能执行本宏命令！")
            mycheck = False
        End If
    Else
        MsgBox ("文档中没有表格，不能执行本宏命令！")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '读取：表格类型
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '读取：姓名或单位名称
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '读取：申请日期
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '读取：内容描述
    
    CH = IIf(Left(bh, 1) = "G", "你", "你单位")
    T0 = "回执第" & bh & "号"
    t1 = "通过电子邮件方式提出政府信息公开申请，申请获得“" & bt & "”信息。"
    t2 = "经查，" & CH & "的申请行为符合《中华人民共和国政府信息公开条例》第二十条规定，我局予以受理。" & Chr(13)
    t3 = "　　根据《中华人民共和国政府信息公开条例》第二十四条，对" & CH & "的申请，我局将于"
    t4 = "前作出书面答复。" & Chr(13) & "　　特此告知。"
    t5 = "国家统计局统计资料管理中心" & Chr(13)
    
    dt1 = Year(dt) & "年" & Month(dt) & "月" & Day(dt) & "日"
    
    td = Date
    fd = DateAdd("d", 21, td)
    dt2 = Year(fd) & "年" & Month(fd) & "月" & Day(fd) & "日"
    dt3 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    Documents.Open ("D:\信息公开\登记回执.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "："
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="　　" & dt1 & "，" & CH & t1 & t2 & t3 & dt2 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="回执" & bh, FileFormat:=wdFormatDocument    '保存文件
    
    Exit Do
    Loop

End Sub

'-------------------------------------------------------------------------
'本宏命令的功能：读取申请表信息、自动生成部分公开告知书并保存到指定目录
'作者：戴宏国
'日期：2008年5月8日
'-------------------------------------------------------------------------

Sub G03_部告()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Set MyDOC = Application.ActiveWindow.Document  '确定活动文档为工作文档
    If MyDOC.Tables.Count > 0 Then  '判断文档中是否有表格
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        If TN = "国家统计局政府信息公开申请表" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
            '判断是否已经登记
            If Len(bh) < 10 Then
                MsgBox ("申请表尚未登记，请先登记！")
                mycheck = False
            Else
                '判断是否已经生成过登记回执
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\信息公开"
                        .FileName = "部告" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("已经生成过登记回执，不能再次生成！")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("非政府信息公开申请表，不能执行本宏命令！")
            mycheck = False
        End If
    Else
        MsgBox ("文档中没有表格，不能执行本宏命令！")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '读取：表格类型
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '读取：姓名或单位名称
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '读取：申请日期
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '读取：内容描述
    
    CH = IIf(Left(bh, 1) = "G", "你", "你单位")
    T0 = "部告第" & bh & "号"
    t1 = "我局受理了" & CH & "提出的政府信息公开申请，具体见《登记回执》第" & bh & "号。" & Chr(13)
    t2 = "　　经查，" & CH & "申请获取的信息属于部分公开范围。根据《中华人民共和国政府信息公开条例》第二十二条规定，我局将以电子邮件方式提供可以公开部分的政府信息。" & Chr(13)
    t3 = "　　" & CH & "申请获取的政府信息中，有部分内容属于：" & Chr(11) & "　　● 申请的统计数据需要进行再加工、处理" & Chr(13)
    t3 = t3 & "　　● 国家机密" & Chr(13) & "　　● 商业机密或者公开可能导致商业机密被泄露的政府信息" & Chr(13)
    t3 = t3 & "　　● 属于个人隐私或者公开可能导致对个人隐私权造成不当侵害的政府信息" & Chr(13) & "　　● 法律、法规规定不予公开的其他情形。" & Chr(13)
    t4 = "　　根据《中华人民共和国政府信息公开条例》第十四条，对于" & CH & "申请获取的部分信息，我局不予公开。" & Chr(13) & "　　特此告知。"
    t5 = "国家统计局统计资料管理中心" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "年" & sl_m & "月" & sl_d & "日"
    dt2 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) + 21 & "日"
    dt3 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    Documents.Open ("D:\信息公开\G02_部分公开告知书.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "："
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="　　" & dt1 & "，" & t1 & t2 & t3 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="部告" & bh, FileFormat:=wdFormatDocument    '保存文件
    
    Exit Do
    Loop

End Sub

'-------------------------------------------------------------------------------
'本宏命令的功能：读取申请表信息、自动生成政府信息不予公开告知书并保存到指定目录
'作者：戴宏国
'日期：2008年5月8日
'-------------------------------------------------------------------------------

Sub G04_不告()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Set MyDOC = Application.ActiveWindow.Document  '确定活动文档为工作文档
    If MyDOC.Tables.Count > 0 Then  '判断文档中是否有表格
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        If TN = "国家统计局政府信息公开申请表" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
            '判断是否已经登记
            If Len(bh) < 10 Then
                MsgBox ("申请表尚未登记，请先登记！")
                mycheck = False
            Else
                '判断是否已经生成过登记回执
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\信息公开"
                        .FileName = "不告" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("已经生成过不予公开告知书，不能再次生成！")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("非政府信息公开申请表，不能执行本宏命令！")
            mycheck = False
        End If
    Else
        MsgBox ("文档中没有表格，不能执行本宏命令！")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '读取：表格类型
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '读取：姓名或单位名称
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '读取：申请日期
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '读取：内容描述
    
    CH = IIf(Left(bh, 1) = "G", "你", "你单位")
    T0 = "不告第" & bh & "号"
    t1 = "我局受理了" & CH & "提出的政府信息公开申请，具体见《登记回执》第" & bh & "号。"
    t2 = "经查，" & CH & "申请获取的信息属于：" & Chr(13)
    t3 = "　　● 申请的统计数据需要进行再加工、处理" & Chr(13)
    t3 = t3 & "　　● 国家机密" & Chr(13) & "　　● 商业机密或者公开可能导致商业机密被泄露的政府信息" & Chr(13)
    t3 = t3 & "　　● 属于个人隐私或者公开可能导致对个人隐私权造成不当侵害的政府信息" & Chr(13) & "　　● 法律、法规规定不予公开的其他情形。" & Chr(13)
    t4 = "　　根据《中华人民共和国政府信息公开条例》第二十一条第二款，对于" & CH & "申请获取的信息，我局不予公开。" & Chr(13) & "　　特此告知。"
    t5 = "国家统计局统计资料管理中心" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "年" & sl_m & "月" & sl_d & "日"
    dt2 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) + 21 & "日"
    dt3 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    Documents.Open ("D:\信息公开\G03_不予公开告知书.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "："
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="　　" & dt1 & "，" & t1 & t2 & t3 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="不告" & bh, FileFormat:=wdFormatDocument    '保存文件
    
    Exit Do
    Loop

End Sub

'-----------------------------------------------------------------------------
'本宏命令的功能：读取申请表信息、自动生成政府信息不存在告知书并保存到指定目录
'作者：戴宏国
'日期：2008年5月8日
'-----------------------------------------------------------------------------

Sub G05_不存告()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Set MyDOC = Application.ActiveWindow.Document  '确定活动文档为工作文档
    If MyDOC.Tables.Count > 0 Then  '判断文档中是否有表格
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        If TN = "国家统计局政府信息公开申请表" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
            '判断是否已经登记
            If Len(bh) < 10 Then
                MsgBox ("申请表尚未登记，请先登记！")
                mycheck = False
            Else
                '判断是否已经生成过登记回执
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\信息公开"
                        .FileName = "不存告" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("已经生成过信息不存在告知书，不能再次生成！")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("非政府信息公开申请表，不能执行本宏命令！")
            mycheck = False
        End If
    Else
        MsgBox ("文档中没有表格，不能执行本宏命令！")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '读取：表格类型
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '读取：姓名或单位名称
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '读取：申请日期
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '读取：内容描述
    
    CH = IIf(Left(bh, 1) = "G", "你", "你单位")
    T0 = "不存告第" & bh & "号"
    t1 = "我局受理了" & CH & "提出的政府信息公开申请，具体见《登记回执》第" & bh & "号。" & Chr(13)
    t2 = "　　经查，" & CH & "申请获取的信息不存在。" & Chr(13)
'    T3 = "　　根据《中华人民共和国政府信息公开条例》第二十四条，对" & CH & "的申请，我局将于"
    t4 = "　　特此告知。"
    t5 = "国家统计局统计资料管理中心" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "年" & sl_m & "月" & sl_d & "日"
    dt2 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) + 21 & "日"
    dt3 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    Documents.Open ("D:\信息公开\G04_信息不存在告知书.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "："
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="　　" & dt1 & "，" & t1 & t2 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="不存告" & bh, FileFormat:=wdFormatDocument    '保存文件
    
    Exit Do
    Loop

End Sub

'-------------------------------------------------------------------------------------
'本宏命令的功能：读取申请表信息、自动生成非本机关政府信息公开告知书并保存到指定目录
'作者：戴宏国
'日期：2008年5月8日
'-------------------------------------------------------------------------------------

Sub G06_非告()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Set MyDOC = Application.ActiveWindow.Document  '确定活动文档为工作文档
    If MyDOC.Tables.Count > 0 Then  '判断文档中是否有表格
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        If TN = "国家统计局政府信息公开申请表" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
            '判断是否已经登记
            If Len(bh) < 10 Then
                MsgBox ("申请表尚未登记，请先登记！")
                mycheck = False
            Else
                '判断是否已经生成过登记回执
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\信息公开"
                        .FileName = "非告" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("已经生成过非本机关信息知书，不能再次生成！")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("非政府信息公开申请表，不能执行本宏命令！")
            mycheck = False
        End If
    Else
        MsgBox ("文档中没有表格，不能执行本宏命令！")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '读取：表格类型
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '读取：姓名或单位名称
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '读取：申请日期
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '读取：内容描述
    
    CH = IIf(Left(bh, 1) = "G", "你", "你单位")
    T0 = "非告第" & bh & "号"
    t1 = "我局受理了" & CH & "提出的政府信息公开申请，具体见《登记回执》第" & bh & "号。" & Chr(13)
    t2 = "　　经查，" & CH & "申请获取的信息不属于本机关的掌握范围，建议向____机关咨询，联系方式为____。" & Chr(13)
'    T3 = "　　根据《中华人民共和国政府信息公开条例》第二十四条，对" & CH & "的申请，我局将于"
    t4 = "　　特此告知。"
    t5 = "国家统计局统计资料管理中心" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "年" & sl_m & "月" & sl_d & "日"
    dt2 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) + 21 & "日"
    dt3 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    Documents.Open ("D:\信息公开\G05_非本机关政府信息告知书.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "："
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="　　" & dt1 & "，" & t1 & t2 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="非告" & bh, FileFormat:=wdFormatDocument    '保存文件
    
    Exit Do
    Loop

End Sub

'-------------------------------------------------------------------------
'本宏命令的功能：读取申请表信息、自动生成补正申请通知书并保存到指定目录
'作者：戴宏国
'日期：2008年5月8日
'-------------------------------------------------------------------------

Sub G07_补通()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Set MyDOC = Application.ActiveWindow.Document  '确定活动文档为工作文档
    If MyDOC.Tables.Count > 0 Then  '判断文档中是否有表格
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        If TN = "国家统计局政府信息公开申请表" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
            '判断是否已经登记
            If Len(bh) < 10 Then
                MsgBox ("申请表尚未登记，请先登记！")
                mycheck = False
            Else
                '判断是否已经生成过登记回执
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\信息公开"
                        .FileName = "补通" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("已经生成过补正申请通知书，不能再次生成！")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("非政府信息公开申请表，不能执行本宏命令！")
            mycheck = False
        End If
    Else
        MsgBox ("文档中没有表格，不能执行本宏命令！")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '读取：表格类型
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '读取：姓名或单位名称
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '读取：申请日期
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '读取：内容描述
    
    CH = IIf(Left(bh, 1) = "G", "你", "你单位")
    T0 = "补通第" & bh & "号"
    t1 = "我局受理了" & CH & "提出的政府信息公开申请，具体见《登记回执》第" & bh & "号。" & Chr(13)
    t2 = "　　经查，" & CH & "申请获取的信息内容不明确，我局难以根据此申请确定具体的政府信息。请更改、补充所需内容描述后再行申请。" & Chr(13)
'    T3 = "　　根据《中华人民共和国政府信息公开条例》第二十四条，对" & CH & "的申请，我局将于"
    t4 = "　　特此告知。"
    t5 = "国家统计局统计资料管理中心" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "年" & sl_m & "月" & sl_d & "日"
    dt2 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) + 21 & "日"
    dt3 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    Documents.Open ("D:\信息公开\G06_补正申请通知书.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "："
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="　　" & dt1 & "，" & t1 & t2 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="补通" & bh, FileFormat:=wdFormatDocument    '保存文件
    
    Exit Do
    Loop

End Sub

'-------------------------------------------------------------------------
'本宏命令的功能：读取申请表信息、自动生成补正申请通知书并保存到指定目录
'作者：戴宏国
'日期：2008年5月8日
'-------------------------------------------------------------------------

Sub G08_告知()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Set MyDOC = Application.ActiveWindow.Document  '确定活动文档为工作文档
    If MyDOC.Tables.Count > 0 Then  '判断文档中是否有表格
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        If TN = "国家统计局政府信息公开申请表" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
            '判断是否已经登记
            If Len(bh) < 10 Then
                MsgBox ("申请表尚未登记，请先登记！")
                mycheck = False
            Else
                '判断是否已经生成过登记回执
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\信息公开"
                        .FileName = "告知" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("已经生成过公开告知书，不能再次生成！")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("非政府信息公开申请表，不能执行本宏命令！")
            mycheck = False
        End If
    Else
        MsgBox ("文档中没有表格，不能执行本宏命令！")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '读取：表格类型
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '读取：姓名或单位名称
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '读取：申请日期
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '读取：内容描述
    
    CH = IIf(Left(bh, 1) = "G", "你", "你单位")
    T0 = "告知第" & bh & "号"
    t1 = "我局受理了" & CH & "提出的政府信息公开申请，具体见《登记回执》第" & bh & "号。" & Chr(13)
    t2 = "　　经查，" & CH & "申请获取的信息属于公开范围。根据《中华人民共和国政府信息公开条例》第二十一条第一款，"
    t3 = "我局将以电子邮件方式提供所申请的政府信息。" & Chr(13)
    t4 = "　　特此告知。"
    t5 = "国家统计局统计资料管理中心" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "年" & sl_m & "月" & sl_d & "日"
    dt2 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) + 21 & "日"
    dt3 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    Documents.Open ("D:\信息公开\G01_公开告知书.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "："
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="　　" & dt1 & "，" & t1 & t2 & t3 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="告知" & bh, FileFormat:=wdFormatDocument    '保存文件
    
    Exit Do
    Loop

End Sub


'-------------------------------------------------------------------------
'本宏命令的功能：读取申请表信息、自动生成补正申请通知书并保存到指定目录
'作者：戴宏国
'日期：2008年5月8日
'-------------------------------------------------------------------------

Sub G09_协办()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Set MyDOC = Application.ActiveWindow.Document  '确定活动文档为工作文档
    If MyDOC.Tables.Count > 0 Then  '判断文档中是否有表格
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        If TN = "国家统计局政府信息公开申请表" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
            '判断是否已经登记
            If Len(bh) < 10 Then
                MsgBox ("申请表尚未登记，请先登记！")
                mycheck = False
            Else
                '判断是否已经生成过登记回执
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\信息公开"
                        .FileName = "协办" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("已经生成过协办通知书，不能再次生成！")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("非政府信息公开申请表，不能执行本宏命令！")
            mycheck = False
        End If
    Else
        MsgBox ("文档中没有表格，不能执行本宏命令！")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '读取：表格类型
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '读取：姓名或单位名称
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '读取：申请日期
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '读取：内容描述
        yt = Left(mytable.Cell(17, 3).Range.Text, Len(mytable.Cell(17, 3).Range.Text) - 2) '读取：用途
        dh = Left(mytable.Cell(9, 3).Range.Text, Len(mytable.Cell(9, 3).Range.Text) - 2) '读取：联系电话
        em = Left(mytable.Cell(11, 3).Range.Text, Len(mytable.Cell(11, 3).Range.Text) - 2) '读取：email
    
    t3 = "　　现将我局已登记受理的第" & bh
    t3 = t3 & "号政府信息公开申请转发给贵单位，请协助办理，期限为十个工作日。"
    t3 = t3 & "请在办理完毕后尽快将办理结果发给我们，再由我们统一答复申请人。谢谢！" & Chr(13)
    t4 = "　　特此通知。"
    t5 = "国家统计局统计资料管理中心" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "年" & sl_m & "月" & sl_d & "日"
    dt2 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) + 21 & "日"
    dt3 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    
    Documents.Open ("D:\信息公开\G07_协办通知.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=t3 & t4
        mt.Cell(5, 2).Select
        Selection.TypeText Text:=t5 & dt3
    Set st = ActiveDocument.Tables(2)
        st.Cell(2, 2).Select
        Selection.TypeText Text:=bh
        st.Cell(3, 2).Select
        Selection.TypeText Text:=dt1
        st.Cell(4, 2).Select
        Selection.TypeText Text:=xm
        st.Cell(5, 2).Select
        Selection.TypeText Text:=dh
        st.Cell(6, 2).Select
        Selection.TypeText Text:=em
        st.Cell(7, 2).Select
        Selection.TypeText Text:=dt
        st.Cell(8, 2).Select
        Selection.TypeText Text:=bt
        st.Cell(9, 2).Select
        Selection.TypeText Text:=yt
    
    ActiveDocument.SaveAs FileName:="协办" & bh, FileFormat:=wdFormatDocument    '保存文件
    
    Exit Do
    Loop

End Sub

'-------------------------------------------------------------------------
'本宏命令的功能：读取申请表信息、自动生成补正申请通知书并保存到指定目录
'作者：戴宏国
'日期：2008年5月8日
'-------------------------------------------------------------------------

Sub G10_协调()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Set MyDOC = Application.ActiveWindow.Document  '确定活动文档为工作文档
    If MyDOC.Tables.Count > 0 Then  '判断文档中是否有表格
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        If TN = "国家统计局政府信息公开申请表" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
            '判断是否已经登记
            If Len(bh) < 10 Then
                MsgBox ("申请表尚未登记，请先登记！")
                mycheck = False
            Else
                '判断是否已经生成过登记回执
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\信息公开"
                        .FileName = "协调" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("已经生成过协调通知书，不能再次生成！")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("非政府信息公开申请表，不能执行本宏命令！")
            mycheck = False
        End If
    Else
        MsgBox ("文档中没有表格，不能执行本宏命令！")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '读取：表格类型
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '读取：姓名或单位名称
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '读取：申请日期
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '读取：内容描述
        yt = Left(mytable.Cell(17, 3).Range.Text, Len(mytable.Cell(17, 3).Range.Text) - 2) '读取：用途
        dh = Left(mytable.Cell(9, 3).Range.Text, Len(mytable.Cell(9, 3).Range.Text) - 2) '读取：联系电话
        em = Left(mytable.Cell(11, 3).Range.Text, Len(mytable.Cell(11, 3).Range.Text) - 2) '读取：email
    
    t3 = "　　现将我局已登记受理的第" & bh
    t3 = t3 & "号政府信息公开申请转发给贵单位，请协调有关司级别单位办理，期限为十个工作日。"
    t3 = t3 & "请在主办单位在办理完毕后尽快将办理结果发给我们，再由我们统一答复申请人。谢谢！" & Chr(13)
    t4 = "　　特此通知。"
    t5 = "国家统计局统计资料管理中心" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "年" & sl_m & "月" & sl_d & "日"
    dt2 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) + 21 & "日"
    dt3 = Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    
    Documents.Open ("D:\信息公开\G08_协调通知.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=t3 & t4
        mt.Cell(5, 2).Select
        Selection.TypeText Text:=t5 & dt3
    Set st = ActiveDocument.Tables(2)
        st.Cell(2, 2).Select
        Selection.TypeText Text:=bh
        st.Cell(3, 2).Select
        Selection.TypeText Text:=dt1
        st.Cell(4, 2).Select
        Selection.TypeText Text:=xm
        st.Cell(5, 2).Select
        Selection.TypeText Text:=dh
        st.Cell(6, 2).Select
        Selection.TypeText Text:=em
        st.Cell(7, 2).Select
        Selection.TypeText Text:=dt
        st.Cell(8, 2).Select
        Selection.TypeText Text:=bt
        st.Cell(9, 2).Select
        Selection.TypeText Text:=yt
    
    ActiveDocument.SaveAs FileName:="协调" & bh, FileFormat:=wdFormatDocument    '保存文件
    
    Exit Do
    Loop

End Sub


'----------------------------------------------------------------------------
'本宏命令的功能：自动按不同情况对申请表进行编号、保存、登记
'作者：戴宏国
'日期：2008年5月5日
'----------------------------------------------------------------------------
Sub E01_登记()

    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Set MyDOC = Application.ActiveWindow.Document  '确定活动文档为工作文档
    
    If MyDOC.Tables.Count = 0 Then  '判断文档中是否有表格
        MsgBox ("文档中没有表格，不能执行本宏命令！")
        mycheck = False
    Else
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        MsgBox TN
        If TN <> "Application Form for the Publication of Government Information of NBS" Then
            MsgBox ("非政府信息公开申请表，不能执行本宏命令！")
            mycheck = False
        Else
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
            lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '读取表格类型
            dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '读取：申请日期
            '判断是否已经登记
            If Len(bh) < 10 Then
                If IsDate(dt) Then
                    mycheck = True
                Else
                    MsgBox "申请日期格式有误，请先修改，再执行本命令！" & Chr(13) & "正确的格式为：2008-5-1，2008-05-01或2008年5月1日"
                    mycheck = False
                End If
            Else
                MsgBox ("此申请表已登记，不能重复登记！")
                mycheck = False
            End If
        End If
    End If

    Do While mycheck = True
    
    Application.DisplayAlerts = wdAlertsNone
    M = IIf(Month(Date) < 10, "0" & Month(Date), Month(Date))
    D = IIf(Day(Date) < 10, "0" & Day(Date), Day(Date))
    y = Year(Date)
    If lb = "For Individual" Then
    N = "G" & y & M & D & "-"
    Else
    N = "D" & y & M & D & "-"
    End If

'在工作目录中查找是否有当天存入的文件，如果没有，编号从1开始，否则，在原来的编号基础上连续编号
Set fs = Application.FileSearch
With fs
    .LookIn = "D:\信息公开"
    .FileName = "申请" & N
    
    If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
        FN = N & 1
    Else
        FN = N & .FoundFiles.Count + 1
    End If
End With

    If Len(ActiveDocument.Tables(1).Cell(2, 2).Range.Text) < Len(FN) Then
    ActiveDocument.Tables(1).Cell(2, 2).Select
    Selection.TypeText Text:=FN
    ActiveDocument.SaveAs FileName:="申请" & FN, FileFormat:=wdFormatDocument   '保存文件
    End If
    
    If lb = "For Individual" Then
    Set MyDOC = ActiveDocument
    Set MyTab = MyDOC.Tables(1)
    t1 = Left(MyTab.Cell(5, 3).Range.Text, Len(MyTab.Cell(5, 3).Range.Text) - 2)
    t2 = Left(MyTab.Cell(6, 3).Range.Text, Len(MyTab.Cell(6, 3).Range.Text) - 2)
    t3 = Left(MyTab.Cell(7, 3).Range.Text, Len(MyTab.Cell(7, 3).Range.Text) - 2)
    t4 = Left(MyTab.Cell(8, 3).Range.Text, Len(MyTab.Cell(8, 3).Range.Text) - 2)
    t5 = Left(MyTab.Cell(9, 3).Range.Text, Len(MyTab.Cell(9, 3).Range.Text) - 2)
    t6 = Left(MyTab.Cell(10, 3).Range.Text, Len(MyTab.Cell(10, 3).Range.Text) - 2)
    t7 = Left(MyTab.Cell(11, 3).Range.Text, Len(MyTab.Cell(11, 3).Range.Text) - 2)
    t8 = Left(MyTab.Cell(12, 3).Range.Text, Len(MyTab.Cell(12, 3).Range.Text) - 2)
    t9 = Left(MyTab.Cell(13, 3).Range.Text, Len(MyTab.Cell(13, 3).Range.Text) - 2)
    t10 = Left(MyTab.Cell(14, 3).Range.Text, Len(MyTab.Cell(14, 3).Range.Text) - 2)
    t11 = Left(MyTab.Cell(16, 3).Range.Text, Len(MyTab.Cell(16, 3).Range.Text) - 2)
    t12 = Left(MyTab.Cell(17, 3).Range.Text, Len(MyTab.Cell(17, 3).Range.Text) - 2)
    t13 = Left(MyTab.Cell(18, 3).Range.Text, Len(MyTab.Cell(18, 3).Range.Text) - 2)
    t14 = Left(MyTab.Cell(19, 3).Range.Text, Len(MyTab.Cell(19, 3).Range.Text) - 2)
'    t15 = Left(mytab.Cell(20, 3).Range.Text, Len(mytab.Cell(20, 3).Range.Text) - 2)

    Documents.Open FileName:="D:\信息公开\登记表3.doc"
    Set MyDOC = ActiveDocument
    Set MyTab = MyDOC.Tables(1)
    N = MyTab.Rows.Count
    i = 0
    MyTab.Cell(N, 1).Select
    Selection.InsertRowsBelow 1

    A = Array(N, FN, "", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14)
    
    For Each aCell In MyTab.Rows(N + 1).Range.Cells
        aCell.Range.Text = A(i)
        i = i + 1
    Next aCell
    ActiveDocument.Close SaveChanges:=wdSaveChanges
'    MsgBox ("已成功登记！")
    End If
    
    If lb = "For Organization" Then
    Set MyDOC = ActiveDocument
    Set MyTab = MyDOC.Tables(1)
    t1 = Left(MyTab.Cell(5, 3).Range.Text, Len(MyTab.Cell(5, 3).Range.Text) - 2)
    t2 = Left(MyTab.Cell(6, 3).Range.Text, Len(MyTab.Cell(6, 3).Range.Text) - 2)
    t3 = Left(MyTab.Cell(7, 3).Range.Text, Len(MyTab.Cell(7, 3).Range.Text) - 2)
    t4 = Left(MyTab.Cell(8, 3).Range.Text, Len(MyTab.Cell(8, 3).Range.Text) - 2)
    t5 = Left(MyTab.Cell(9, 3).Range.Text, Len(MyTab.Cell(9, 3).Range.Text) - 2)
    t6 = Left(MyTab.Cell(10, 3).Range.Text, Len(MyTab.Cell(10, 3).Range.Text) - 2)
    t7 = Left(MyTab.Cell(11, 3).Range.Text, Len(MyTab.Cell(11, 3).Range.Text) - 2)
    t8 = Left(MyTab.Cell(12, 3).Range.Text, Len(MyTab.Cell(12, 3).Range.Text) - 2)
    t9 = Left(MyTab.Cell(13, 3).Range.Text, Len(MyTab.Cell(13, 3).Range.Text) - 2)
    t10 = Left(MyTab.Cell(14, 3).Range.Text, Len(MyTab.Cell(14, 3).Range.Text) - 2)
    t11 = Left(MyTab.Cell(16, 3).Range.Text, Len(MyTab.Cell(16, 3).Range.Text) - 2)
    t12 = Left(MyTab.Cell(17, 3).Range.Text, Len(MyTab.Cell(17, 3).Range.Text) - 2)
    t13 = Left(MyTab.Cell(18, 3).Range.Text, Len(MyTab.Cell(18, 3).Range.Text) - 2)
    t14 = Left(MyTab.Cell(19, 3).Range.Text, Len(MyTab.Cell(19, 3).Range.Text) - 2)

    Documents.Open FileName:="D:\信息公开\登记表4.doc"
    Set MyDOC = ActiveDocument
    Set MyTab = MyDOC.Tables(1)
    N = MyTab.Rows.Count
    i = 0
    MyTab.Cell(N, 1).Select
    Selection.InsertRowsBelow 1

    A = Array(N, FN, "", t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14)
    
    For Each aCell In MyTab.Rows(N + 1).Range.Cells
        aCell.Range.Text = A(i)
        i = i + 1
    Next aCell
    ActiveDocument.Close SaveChanges:=wdSaveChanges
'    MsgBox ("已成功登记！")
    End If
    
    Exit Do
    Loop
    
End Sub

'--------------------------------------------------------------
'本宏命令的功能：读取申请表信息、自动生成回执并保存到指定目录
'作者：戴宏国
'日期：2008年5月6日
'--------------------------------------------------------------

Sub E02_回执()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Set MyDOC = Application.ActiveWindow.Document  '确定活动文档为工作文档
    If MyDOC.Tables.Count > 0 Then  '判断文档中是否有表格
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        MsgBox TN
        If TN = "Application Form for the Publication of Government Information of NBS" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
            '判断是否已经登记
            If Len(bh) < 10 Then
                MsgBox ("申请表尚未登记，请先登记！")
                mycheck = False
            Else
                '判断是否已经生成过登记回执
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\信息公开"
                        .FileName = "回执" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("已经生成过登记回执，不能再次生成！")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("非政府信息公开申请表，不能执行本宏命令！")
            mycheck = False
        End If
    Else
        MsgBox ("文档中没有表格，不能执行本宏命令！")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '读取：表格名称
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '读取：表格类型
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '读取：编号
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '读取：姓名或单位名称
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '读取：申请日期
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '读取：内容描述
    
'    CH = IIf(Left(bh, 1) = "G", "你", "你单位")
    T0 = "Feedback No.: " & bh
    D = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    dt1 = Day(dt) & " " & D(Month(dt) - 1) & " " & Year(dt) & ","
    
    t1 = Space(4) & dt1 & " you applied for the information on " & bt & " through e-mail. "
    t1 = t1 & "According to the related Articles of the Regulation on "
    t1 = t1 & "the Publication of Government Information of the People's Republic of China, We would like to "
    t1 = t1 & "inform you that your application was accepted and will be replied before "
    t5 = "Statistical Library and Information Services (SLIS)" & Chr(13) & "National Bureau of Statistics of China (NBS)" & Chr(13)
    t6 = Space(4) & "If you have any questions, please do not hesitate to contact us." & Chr(13)
    
    td = Date
    fd = DateAdd("d", 21, td)
    dt2 = Day(fd) & " " & D(Month(fd) - 1) & " " & Year(fd)
    dt3 = Day(td) & " " & D(Month(td) - 1) & " " & Year(td)
    
    Documents.Open ("D:\信息公开\E01_Feedback.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:="Mr. /Mrs. " & xm & ","
        mt.Cell(5, 1).Select
        Selection.TypeText Text:=t1 & dt2 & " by NBS." & Chr(13) & t6
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="回执" & bh, FileFormat:=wdFormatDocument    '保存文件
    
    Exit Do
    Loop

End Sub


Function Disp(LN)
            TT = LN
        C = Chr(9)
        L = Len(TT)
        P = InStr(1, TT, C, 1)
        S1 = Trim(Left(TT, P - 1))
        S2 = Right(TT, L - P)
        x1 = S1
        TT = S2
        L = Len(TT)
        P = InStr(1, TT, C, 1)
        S1 = Trim(Left(TT, P - 1))
        S2 = Right(TT, L - P)
        x2 = S1
        TT = S2
        L = Len(TT)
        P = InStr(1, TT, C, 1)
        S1 = Trim(Left(TT, P - 1))
        S2 = Right(TT, L - P)
        x3 = S1
        TT = S2
        L = Len(TT)
        P = InStr(1, TT, C, 1)
        S1 = Trim(Left(TT, P - 1))
        S2 = Right(TT, L - P)
        x4 = S1
        TT = S2
        L = Len(TT)
        P = InStr(1, TT, C, 1)
        S1 = Trim(Left(TT, P - 1))
        S2 = Right(TT, L - P)
        x5 = S1
        x6 = S2
        
        MsgBox "姓名：" & x1 & Chr(13) & "单位：" & x2 & " " & Chr(13) & "职务：" & x3 & Chr(13) & "分机：" & x4 & Chr(13) & "直拨机：" & x5 & Chr(13) & "电邮：" & x6 & Chr(13)
        
 '       TextBox3.Text = x1
 '       TextBox4.Text = x2 & " " & x3
 '       TextBox5.Text = x4
 '       TextBox6.Text = x5
 '       If Len(x6) > 0 Then
 '       TextBox7.Text = x6 & "@stats.gov.cn"
 '       Else
 '       TextBox7.Text = x6
 '       End If
        

End Function



'---------------------------------------------------
'本宏命令的功能：自动将申请表编号并保存到指定目录
'作者：戴宏国
'日期：2008年5月7日
'---------------------------------------------------

Sub B01_测试()

    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    fl = "NBSTEL.TXT"    ' 需要转换的文件名
    DL = Chr(9)       ' 分隔符
    TXT = ""            ' 声明一个空字符串变量
    NM = InputBox("请输入你要查询的内容：")
    
    Open fl For Input As #1     ' 打开输入文件。
    MsgBox LOF(1)
    Do While Not EOF(1)         ' 循环至文件尾。
        Line Input #1, LN       ' 把第一行字符串赋值给变量LN
        If InStr(1, LN, NM, 1) Then
'            MsgBox LN
        L = Split(LN, delimiter:=DL) ' 按分隔符把字符串赋值给数组L
        MsgBox "姓名：" & L(0) & Chr(13) & "单位：" & L(1) & " " & Chr(13) & "职务：" & L(2) & Chr(13) & "分机：" & L(3) & Chr(13) & "直拨：" & L(4) & Chr(13) & "电邮：" & L(5) & Chr(13), , "查询结果"
        End If
        
'        L = Split(LN, delimiter:=DL) ' 按分隔符把字符串赋值给数组L
'        For i = 0 To UBound(L)
'        txt = txt & Chr(9) & Trim(L(i))
'        Next i
'        txt = Right(txt, Len(txt) - 1)
'        Selection.TypeText Text:=txt & Chr(13)
    Loop        ' 循环
            L = Loc(1)
            MsgBox Loc(1)

    Close #1    ' 关闭文件

    Open fl For Output As #1     ' 打开输入文件。

Seek #1, L
Write #1, Chr(13) & "This is a test!" & Chr(13)
 Close #1    ' 关闭文件

End Sub


'---------------------------------------------------
'本宏命令的功能：自动将申请表编号并保存到指定目录
'作者：戴宏国
'日期：2008年5月27日
'---------------------------------------------------

Sub 自动查拼音()
    ChangeFileOpenDirectory "D:\信息公开"      '确定工作目录
    Dim Z() As Variant, P() As Variant, Q() As Variant
    
    fl = "pinyin.doc"    ' 拼音对照表文件名
    DL = Chr(9)       ' 分隔符
    txt1 = ""            ' 声明一个空字符串变量
    txt2 = ""
    Do While True
    NM = InputBox("请输入你要查询的内容：")
    If Len(NM) > 1 And Asc(Left(NM, 1)) < 0 Then
    
    Exit Do
    End If
    Loop
    
    Set MyDoc1 = Documents.Add
    Selection.TypeText Text:=NM
    Selection.Paragraphs(1).Range.Select
    N = Selection.Characters.Count
    For i = 1 To N - 1
    ReDim Preserve Z(i)
    Z(i) = Selection.Characters(i)
    Next i
    MyDoc1.Close SaveChanges:=wdDoNotSaveChanges

    
    Documents.Open FileName:=fl, Visible:=False   ' 打开拼音对照表文件
'    Set Mydoc2 = ActiveDocument
    Documents(fl).Activate
    
    For i = 1 To UBound(Z())
    ZZ = Z(i)
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=ZZ
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
            TT = Selection.Paragraphs(1).Range.Text
            FF = Split(TT, delimiter:=DL)
            ReDim Preserve P(i)
            ReDim Preserve Q(i)
            P(i) = Left(FF(0), 1)
            txt1 = txt1 & P(i)
            Q(i) = FF(0)
            txt2 = txt2 & Q(i) & " "
        End If
        Selection.WholeStory
        Selection.HomeKey
    End With
    Next i
    Documents(fl).Close
    
    MsgBox txt1 & "  " & txt2
End Sub

Sub 文件对话框()
    Dim fd As FileDialog    '声明一个文件对话框变量
    Set fd = Application.FileDialog(msoFileDialogFilePicker)    '创建文件对话框对象
    Dim fs As Variant
    With fd
        .AllowMultiSelect = False   '只允许选择一个文件
        If .Show = -1 Then  '用show方法来显示文件选择对话框
            For Each fs In .SelectedItems   '遍历FileDialogSelectedItems集合每个成员
                MsgBox "选择的文件路径为: " & fs
            Next fs
        Else    '用户按了取消键
        End If
    End With
    Set fd = Nothing    '释放定义的对象

End Sub

Sub 查找次数()
     Do While True
    NM = InputBox("请输入你要查询的内容：")
    If Len(NM) > 0 Then
    
    Exit Do
    End If
    Loop
    Set MyDoc1 = ActiveDocument
        
        fl = "C:\EasyQuery\NBSTEL.doc"
        Documents.Open FileName:=fl, Visible:=False   ' 打开文件
        Documents(fl).Activate
        Selection.WholeStory
        Selection.HomeKey Unit:=wdStory
    
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = NM
        .Forward = True
        .Wrap = wdFindContinue
    End With
    
    i = 0
    Do While True
    Selection.Find.Execute
    If Selection.Find.Found = True Then
        i = i + 1
    Else
        Exit Do
    End If
    Loop
    
    If i > 0 Then
        MsgBox "找到 " & i & " 个", , "提示"
    Else
        MsgBox "未找到！", , "提示"
    End If
    
    Documents(fl).Close
    MyDoc1.Activate

End Sub

Function fnd(TXT)
    Dim pn() As Variant
    
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = TXT
        .Forward = True
        .Wrap = wdFindContinue
    End With
    
    i = 0
    Do While True
    Selection.Find.Execute
    If Selection.Find.Found = True Then
        i = i + 1
        ReDim Preserve pn(i)
        Set MyRange = ActiveDocument.Range(start:=0, End:=Selection.Paragraphs(1).Range.End)
        MyRange.Select
        N = Selection.Paragraphs.Count
        pn(i) = N
 '       MsgBox "在第 " & n & " 段找到一个"
        Selection.EndKey
        Selection.MoveRight Unit:=wdCharacter, Count:=1
    Else
        Exit Do
    End If
    Loop
    
    If i > 0 Then
        st = ""
        For j = 1 To UBound(pn())
        st = st & pn(j) & " "
        Next j
        
        MsgBox "共找到 " & i & " 个！" & "所处的段落分别是：" & st, , "提示"
    Else
        MsgBox "未找到！", , "提示"
    End If
End Function

Sub mytest1()
    Dim D As Document
    Dim R As Range
    Dim P As Paragraphs
    
    
    Set D = ActiveDocument
    Set P = D.Paragraphs
    Set R = D.Range(start:=0, End:=D.Paragraphs(3).Range.End)
    
    N = P.Count
    MsgBox N
    
    R.Select
    i = Selection.Paragraphs.Count
    MsgBox i
    
    R.SetRange start:=0, End:=D.Paragraphs(3 + i).Range.End
    R.Select
    i = Selection.Paragraphs.Count
    MsgBox i
    
    P(50).Range.Select
    T = Selection.Range.Text
    MsgBox T
End Sub

    Option Base 1
    Public txt1 As String, txt2 As String, Rest As String
    
Sub pinyin()
    fc = "C:\EasyQuery\导入测试文件.doc"
    Documents.Open FileName:=fc, Visible:=True
    Documents(fc).Activate
    Set my_doc = ActiveDocument
'    If ActiveDocument.Tables.Count > 0 Then
'        Set mytab = ActiveDocument.Tables(1)
'        cn = mytab.Columns.Count
'        rn = mytab.Rows.Count
'        MsgBox "表格列数：" & cn & "  " & "表格列数：" & rn
'        c11 = mytab.Cell(1, 1).Range.Text
'        c11 = Left(c11, Len(c11) - 2)
'        MsgBox c11 & " " & ln1 & "  " & ln2
'    End If
    
Documents.Open FileName:="C:\EasyQuery\pinyin.doc", Visible:=True

'    If cn < 10 Then
'        For j = 1 To 8 - cn
'            mytab.Columns.Add
'        Next j
'    End If
    
'    my_doc.Range.Select
'    Selection.Copy
'    my_doc.Close savechanges:=wdDoNotSaveChanges
'    Documents.Add
'    Selection.PasteSpecial DataType:=wdPasteText
'    Selection.TypeBackspace
'    Set mydoc2 = ActiveDocument
    my_doc.Activate
    N = ActiveDocument.Paragraphs.Count
    For i = 1 To N
        T = ActiveDocument.Paragraphs(i).Range.Text
        T = Left(T, Len(T) - 1)
            Find_Pinyin (T)
            my_doc.Activate
  '          MsgBox areturn
    Next i

End Sub
Function Find_Pinyin(T, t1)
    Dim Z() As Variant, P() As Variant, Q() As Variant
    fname = "C:\EasyQuery\pinyin.doc"    ' 拼音对照表文件名
    txt1 = ""            ' 声明一个空字符串变量
    txt2 = ""
    zd = Split(T, Chr(9))
    aname = zd(0)
    Rest = Right(T, Len(T) - Len(aname))
    If Len(aname) = 2 Then
        ReDim Z(2) As Variant
            Z(1) = Left(aname, 1)
            Z(2) = Right(aname, 1)
    Else
        If Len(aname) = 3 Then
            ReDim Z(3) As Variant
            Z(1) = Left(aname, 1)
            Z(2) = Right(Left(aname, 2), 1)
            Z(3) = Right(aname, 1)
        Else
            If Len(aname) >= 4 Then
            ReDim Z(4) As Variant
            Z(1) = Left(aname, 1)
            Z(2) = Right(Left(aname, 2), 1)
            Z(3) = Right(Left(aname, 3), 1)
            Z(4) = Right(Left(aname, 4), 1)
            End If
        End If
    End If
    Doc_Open (fname) ' 打开或激活拼音对照表文件
    Documents("C:\EasyQuery\pinyin.doc").Activate
    
    For i = 1 To UBound(Z())
    ZZ = Z(i)
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=ZZ
        If .Found = True Then
            TT = Selection.Paragraphs(1).Range.Text
            FF = Split(TT, Chr(9))
            ReDim Preserve P(i)
            ReDim Preserve Q(i)
            P(i) = Left(FF(0), 1)
            txt1 = txt1 & P(i)
            Q(i) = FF(0)
            txt2 = txt2 & Q(i) & " "
        End If
        Selection.WholeStory
        Selection.HomeKey
    End With
    Next i
    t1 = T & Chr(9) & Trim(txt1) & Chr(9) & Trim(txt2) & Chr(13)
End Function
Function Doc_Open(fname) '判断文件是否已经打开，如果已打开，则激活，否则，就打开
    For Each doc In Documents
        If doc.name = fname Then Found = True
    Next doc
    If Found <> True Then
        Documents.Open FileName:=fname, Visible:=False  ' 打开电话数据文件
    Else
        Documents(fname).Activate
    End If
End Function
Function Doc_Close(fname) '判断文件是否已经打开，如果已打开，则关闭
    For Each doc In Documents
        If doc.name = fname Then Found = True
    Next doc
    If Found = True Then Documents(fname).Close
End Function
Sub A01_文件路径()
    Set doc = ActiveDocument
    MsgBox NormalTemplate.Path
    MsgBox doc.Path
    MsgBox Options.DefaultFilePath(wdUserTemplatePath)
    MsgBox Templates(1).FullName
    
    For Each aTemp In Templates
        MsgBox aTemp.FullName
    Next aTemp
End Sub
Function fnd_Next(TXT)
    Selection.WholeStory
    Selection.HomeKey
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = TXT
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute
    If Selection.Find.Found = True Then
        TT = Selection.Paragraphs(1).Range.Text
    Else
'        Show_Info ("没有找到你要查询的内容！请重新输入")
        MsgBox "没有找到你要查询的内容！请重新输入"
        TT = ""
        Selection.WholeStory
        Selection.HomeKey
    End If
End Function

Sub Exist_Check(mycheck)
    A = "430922"
    Selection.WholeStory
    Selection.HomeKey
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = A
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute
    If Selection.Find.Found = True Then
        mycheck = True
    Else
'        Show_Info ("没有找到你要查询的内容！请重新输入")
        MsgBox "没有找到你要查询的内容！请重新输入"
        mycheck = False
    End If
End Sub

Sub Show_Result()
    Set doc = ActiveDocument
    
    Call Exist_Check(mycheck)
    If mycheck = True Then
        TT = Selection.Paragraphs(1).Range.Text
        TT = Left(TT, Len(TT) - 1)
        sp = Split(TT, Chr(9))
        Code1 = sp(0)
        Name1 = sp(1)
        Code2 = Left(Code1, 2) & "0000"
        Code3 = Left(Code1, 4) & "00"
        fnd_Next (Code2)
            TT = Selection.Paragraphs(1).Range.Text
            TT = Left(TT, Len(TT) - 1)
            sp = Split(TT, Chr(9))
            Name2 = sp(1)
        fnd_Next (Code3)
            TT = Selection.Paragraphs(1).Range.Text
            TT = Left(TT, Len(TT) - 1)
            sp = Split(TT, Chr(9))
            Name3 = sp(1)
        MsgBox "你要找的是：" & Name1 & Chr(13) & "    代码为：" & Code1 & Chr(13) & "    隶属于：" & Name2 & Name3
    Else
        MsgBox "没有找到！"
    End If
End Sub
Sub Multi_para()
    HouseCalc 99800, 43100
    Call HouseCalc(380950, 49500)
End Sub

Sub HouseCalc(price As Single, wage As Single)
    If 2.5 * wage <= 0.8 * price Then
        MsgBox "You cannot afford this house."
    Else
        MsgBox "This house is affordable."
    End If
End Sub
Sub Char_Array()
    Dim Z() As Variant
    Set MyDOC = ActiveDocument
    MyDOC.Paragraphs(1).Range.Select
    N = Selection.Characters.Count
    For i = 1 To N - 1
    ReDim Preserve Z(i)
    Z(i) = Selection.Characters(i)
    MsgBox Z(i)
    Next i
'    mydoc.Close savechanges:=wdDoNotSaveChanges

End Sub

Sub A01_检查是否启动了Word()
'Documents("NBSTEL.doc").Close
If Tasks.Exists(name:="Microsoft Word") = True Then
    Set myobject = GetObject("", "Word.Application")
'    MsgBox myobject.Name
    Application.Quit SaveChanges:=wdSaveChanges
    Set myobject = Nothing
End If
'Set Word = CreateObject("word.basic")
'Set msword = CreateObject("word.application")
'Set mydoc = msword.Documents.Open("C:\NBS电话查询\NBSTEL.doc", PasswordDocument:="nbsdhg", WritePasswordDocument:="nbsdhg")
'msword.Visible = True

End Sub

Sub GetWord()
    Dim MyWord As Object    '用于存放'Microsoft Word 引用的变量。
    Dim NoWord As Boolean    '用于最后释放的标记。
    On Error Resume Next    '延迟错误捕获。
    Set MyWord = GetObject(, "Word.Application")
    If Err.Number <> 0 Then NoWord = True
    Err.Clear    '如果发生错误则要清除 Err 对象。
    If NoWord <> True Then
'        MyWord.Application.quit
        MyWord.Visible = False
        MsgBox "Hide Word Successfully!"
    Else
        Set MyWord = CreateObject("Word.Application")
        MyWord.Documents.Open ("D:\00 F2008\各单位电话\资料中心（全）.doc")
        MyWord.Visible = False
        MsgBox "成功启动Word并打开" & MyWord.ActiveDocument.name
    End If
    MsgBox "Will Quit!"
    MyWord.Quit
    Unload Form1
End Sub

Sub A01_CSV() '将带有电子邮件地址的记录导出到CSV文件
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set A = fs.CreateTextFile("C:\EasyQuery\Add.csv", True)
    j = ","
    A.WriteLine ("姓名,部门,职务,业务电话,移动电话,住宅电话,电子邮件地址,办公地点")
    Documents.Open FileName:="C:\EasyQuery\NBSTEL.DOC"
    
    For Each PA In ActiveDocument.Paragraphs
    TT = Trim(PA.Range.Text)
    
    fd = Split(TT, Chr(9))
    If Len(Trim(fd(6))) > 1 Then
        T = fd(0) & j & fd(1) & j & fd(2) & j & fd(3) & j & fd(4) & j & fd(5) & j & fd(6) & j & fd(7)
        A.WriteLine (T)
    End If
    Next PA
    A.Close
    MsgBox "导出成功！"
    Documents("NBSTEL.DOC").Close
End Sub


'---------------------------------------------------
'功能：查询手机属地
'作者：戴宏国
'日期：2008年6月10日
'---------------------------------------------------
Sub A01_Mobile() '手机属地查询
    Dim tel As String, fname As String
    FN = "130 131 132 133 134 135 136 137 138 139 159"
    ChangeFileOpenDirectory "C:\EasyQuery" '确定工作目录
    Do While True
        tel = InputBox("请输入你要查询手机号码：" & Chr(13) & "（至少输入前7位）")
        If Len(tel) > 0 Then
            If Len(tel) > 6 Then
                If Asc(Left(tel, 1)) = 49 Then
                    If InStr(1, FN, Left(tel, 3), 1) > 0 Then
                        chk = True
                        Exit Do
                    Else
                        MsgBox "数据库中没有" & Left(tel, 3) & "开头的号码!"
                        chk = False
                        Exit Do
                    End If
                Else
                    MsgBox "输入不正确！手机号必须以“1”开头，请重新输入！"
                    chk = False
                End If
            Else
                MsgBox "请至少输入手机号的前7位数字，你只输入了" & Len(tel) & "位！"
                chk = False
            End If
        Else
            MsgBox "你没有输入，请至少输入手机号的前7位数字！"
            chk = False
        End If
    Loop
    
    Do While chk = True
    fname = Left(tel, 3) & ".doc"
    Documents.Open FileName:=fname, Visible:=False
    Documents(fname).Activate
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = Left(tel, 7)
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute
    If Selection.Find.Found = True Then
        fd = Split(Selection.Paragraphs(1).Range.Text, Chr(124))
        MsgBox "你查找的号码所属地区为：" & fd(1) & "，是" & fd(2)
    Else
         MsgBox "没有找到你要查询的内容！"
    End If
    Documents(fname).Close
    Exit Do
    Loop
End Sub

Option Base 1

Sub A01_Txt_Import()
    Dim PA() As Variant, pn() As Variant, LN As String, MyDOC As Document
    i = 0
    Open "C:\EasyQuery\Test.txt" For Input As #1    ' 打开输入文件。
    Do While Not EOF(1)    ' 循环至文件尾。
        Line Input #1, LN
        i = i + 1
        ReDim Preserve PA(i)
        PA(i) = LN
    Loop
    Close #1    ' 关闭文件
    j = 0
    For i = 1 To UBound(PA)
        L = PA(i)
        fd = Split(L, Chr(9))
        If UBound(fd) = 7 Then
            Call Find_Pinyin(L, t1)
            j = j + 1
            ReDim Preserve pn(j)
            pn(j) = t1
        End If
    Next i
    Set MyDOC = Documents.Add
    MyDOC.Activate
    页面横
    For j = 1 To UBound(pn)
    Selection.TypeText Text:=pn(j)
    Next j
    Selection.TypeBackspace
    Doc_Close ("C:\EasyQuery\pinyin.doc")

End Sub

Sub A01_表格标准化()
    Dim chk As Boolean '声明一个变量，用于判断文档中是否有表格存在
    Dim MyTab As Table '声明一个表格对象变量
    N = ActiveDocument.Tables.Count '取得表格总数
    If N = 0 Then
        MsgBox "文档中没有表格，不能执行该宏命令！" '提示用户文档中没有表格
        chk = False
    Else
        chk = True
    End If
    On Error Resume Next
    Do While chk = True
    For i = 1 To N
    ActiveDocument.Tables(i).Range.Select
    Set MyTab = Selection.Tables(1)
    With MyTab '设置表格属性
        .Rows.Alignment = wdAlignRowRight '表格内容居右对齐
        .TopPadding = CentimetersToPoints(0) '设置单元格上边距为0
        .BottomPadding = CentimetersToPoints(0) '设置单元格下边距为0
        .LeftPadding = CentimetersToPoints(0) '设置单元格左边距为0
        .RightPadding = CentimetersToPoints(0) '设置单元格右边距为0
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone '清楚表格线
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    End With
        Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter '单元格内容垂直居中

    With Selection.Borders(wdBorderTop) '设置顶线
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderBottom) '设置底线
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Font '设置字体
        .NameFarEast = "宋体"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 10.5
    End With
    With Selection.ParagraphFormat '设置段落格式
        .LeftIndent = CentimetersToPoints(0.1)
        .RightIndent = CentimetersToPoints(0.1)
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 12
        .WordWrap = True
    End With
    MyTab.Rows(1).Select '选定表格的第一个单元格
    Selection.SelectRow '选定第一行
    Selection.Rows.Height = CentimetersToPoints(1#) '设置第一行的行高为1厘米
    

    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectColumn
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectRow
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Tables(1).Select
    
    With Selection.Borders(wdBorderVertical)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderHorizontal)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = wdLineWidth150pt
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderBottom)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = wdLineWidth150pt
        .Color = Options.DefaultBorderColor
    End With
    
    Next i
    Exit Do
    Loop

End Sub

Sub A01_简报()
    Dim MyDOC As Document, MyDir As String
    'MyDir = "D:\简报2008"
    'FN = "公众服务情况反映" & Year(Date) & Chr(45)
    'On Error Resume Next
    'ChangeFileOpenDirectory MyDir
    'Set FS = Application.FileSearch
    'With FS
    '    .LookIn = MyDir
    '    .FileName = FN
    '    If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
    '        N = 1
    '    Else
    '        N = .FoundFiles.Count + 1
    '    End If
    'FN = FN & N
    'End With
    
    Set MyDOC = Documents.Add
    For i = 1 To 10
        Selection.TypeParagraph
    Next i
    MyDOC.Range.Select
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 30
        .Alignment = wdAlignParagraphLeft
    End With
    With Selection.Font
        .NameFarEast = "仿宋"
        .NameAscii = "仿宋"
        .Size = 15
    End With
    MyDOC.Paragraphs(1).Range.Select
    Selection.TypeText Text:="公众服务情况反映"
    MyDOC.Paragraphs(1).Range.Select
    With Selection.ParagraphFormat
        .SpaceBefore = 30
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
    End With
    With Selection.Font
        .NameFarEast = "华文中宋"
        .Size = 36
        .Bold = True
        .Color = wdColorRed
    End With
    
    MyDOC.Paragraphs(2).Range.Select
    Selection.TypeText Text:=Year(Date) & "年第 " & "期（总第  期）"
    MyDOC.Paragraphs(2).Range.Select
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
        .SpaceBefore = 15
        .SpaceAfter = 15
    End With

    MyDOC.Paragraphs(3).Range.Select
    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(14.63), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    Selection.TypeText Text:="国家统计局统计资料管理中心" & vbTab & Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    MyDOC.Paragraphs(3).Range.Select
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
    End With

    MyDOC.Paragraphs(4).Range.Select
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = wdLineWidth300pt
        .Color = wdColorRed
    End With
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(0.35)
        .CharacterUnitFirstLineIndent = 2
    End With

    MyDOC.Paragraphs(5).Range.Select
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="标题"
    MyDOC.Paragraphs(5).Range.Select
        With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(0.35)
        .CharacterUnitFirstLineIndent = 2
    End With
    With Selection.Font
        .NameFarEast = "黑体"
        .NameAscii = "黑体"
        .Size = 15
    End With
    For i = 6 To 11
    MyDOC.Paragraphs(i).Range.Select
    With Selection.ParagraphFormat
        .FirstLineIndent = CentimetersToPoints(0.35)
        .CharacterUnitFirstLineIndent = 2
    End With
    Next i
    MyDOC.Paragraphs(6).Range.Select
    Selection.MoveLeft
    'Selection.Sections(1).Footers(1).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberOutside, FirstPage:=True
   ' MyDoc.SaveAs FileName:=FN, FileFormat:=wdFormatDocument    '保存文件

End Sub

Sub A01_每周舆情()
    Dim MyDOC As Document, MyDir As String
    Set MyDOC = Documents.Add
    For i = 1 To 10
        Selection.TypeParagraph
    Next i
    MyDOC.Range.Select
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 30
        .Alignment = wdAlignParagraphLeft
    End With
    With Selection.Font
        .NameFarEast = "仿宋"
        .NameAscii = "仿宋"
        .Size = 15
    End With
    MyDOC.Paragraphs(1).Range.Select
    Selection.TypeText Text:="每周舆情"
    MyDOC.Paragraphs(1).Range.Select
    With Selection.ParagraphFormat
        .SpaceBefore = 60
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
    End With
    With Selection.Font
        .NameFarEast = "华文行楷"
        .Size = 72
        .Bold = False
        .Color = wdColorRed
    End With
    
    MyDOC.Paragraphs(2).Range.Select
    Selection.TypeText Text:=Year(Date) & "年第 " & "期（总第  期）"
    MyDOC.Paragraphs(2).Range.Select
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
        .SpaceBefore = 15
        .SpaceAfter = 15
    End With

    MyDOC.Paragraphs(3).Range.Select
    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(14.63 _
        ), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    Selection.TypeText Text:="国家统计局统计资料管理中心" & vbTab & Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    MyDOC.Paragraphs(3).Range.Select
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
    End With

    MyDOC.Paragraphs(4).Range.Select
    With Selection.Borders(wdBorderTop)
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = wdLineWidth300pt
        .Color = wdColorRed
    End With
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(0.35)
        .CharacterUnitFirstLineIndent = 2
    End With

    MyDOC.Paragraphs(5).Range.Select
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="标题"
    MyDOC.Paragraphs(5).Range.Select
        With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(0.35)
        .CharacterUnitFirstLineIndent = 2
    End With
    With Selection.Font
        .NameFarEast = "黑体"
        .NameAscii = "黑体"
        .Size = 15
    End With
    For i = 6 To 11
    MyDOC.Paragraphs(i).Range.Select
    With Selection.ParagraphFormat
        .FirstLineIndent = CentimetersToPoints(0.35)
        .CharacterUnitFirstLineIndent = 2
    End With
    Next i
    MyDOC.Paragraphs(6).Range.Select
    Selection.MoveLeft
    'Selection.Sections(1).Footers(1).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberOutside, FirstPage:=True
   ' MyDoc.SaveAs FileName:=FN, FileFormat:=wdFormatDocument    '保存文件

End Sub
Sub Char_To_Array()
    Dim name As String
    Dim NA() As Variant
    name = "欧阳典跃"
    For i = 1 To Len(name)
    ReDim Preserve NA(i)
    NA(i) = Mid(name, i, 1)
    Selection.TypeText Text:=NA(i) & Chr(13)
    Next i
End Sub
Sub Name_Doc()
    Dim fgf As String, s As String
    s = Selection.Range.Text
    fgf = Chr(9)
    TXT = ""
    If Len(s) = 0 Then
        MsgBox "请选择做为标题的内容！"
        chk = False
    Else
        chk = True
    End If
    
    Do While chk = True
    N = InStr(1, s, fgf, 1)
    If N > 0 Then
        FA = Split(s, fgf)
        For i = 0 To UBound(FA)
            TXT = TXT & FA(i)
        Next i
        MsgBox TXT
        Exit Do
    Else
        Exit Do
    End If
    Loop
End Sub

Sub 网页图片链接()
    Dim A As Variant
    Dim B As Variant
    A1 = "<imgsrc="
    A2 = "border=0onclick=zoom(this)onload=attachimg(this,'load')alt=/><br/>"
    A3 = "<br/>"
    A4 = "^p^p"
    A5 = "border=0onclick=zoom(this)onload=attachimg(this,'load')alt=/></td>"
    
    A = Array(Chr(-24159), Chr(34), " ", A1, A2, A3, A4, A5)
    B = Array("^p", "", "", "[img]", "[/img]", "", "^p", "[/img]^p")
    
    For i = 0 To UBound(A)
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = A(i)
        .Replacement.Text = B(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    
    Selection.WholeStory

End Sub

Sub A00_创建指定目录()

NM = "F:\00 F" & Year(Date)
NM1 = NM & "\01 OA文件"
NM2 = NM & "\02 党建工作"
NM3 = NM & "\03 纪检工作"
NM4 = NM & "\04 我的文件"
NM5 = NM & "\05 中德项目"
NM6 = NM & "\06 世行项目"
NM7 = NM & "\07 资料中心"
NM8 = NM & "\08 PDF"
NM9 = NM & "\09 下载"
NM10 = NM & "\10  网络文摘"

Dim D As Variant
    D = Array(NM1, NM2, NM3, NM4, NM5, NM6, NM7, NM8, NM9, NM10)
    
Set fs = CreateObject("Scripting.FileSystemObject")

If fs.FolderExists(NM) = False Then
    Set A = fs.CreateFolder(NM)
End If

For i = 0 To UBound(D)
    If fs.FolderExists(D(i)) = False Then
        Set A = fs.CreateFolder(D(i))
    End If
Next i

End Sub


Sub A01_句号前文本加粗()

    Application.ScreenUpdating = False '关闭屏幕更新
    C = Chr(-24157) '句号
    For Each para In ActiveDocument.Paragraphs
        TT = para.Range.Text
        L = Len(TT)
        P = InStr(1, TT, C, 1)
        If P > 0 Then
            S1 = Left(TT, P - 1)
            S2 = Right(TT, L - P)
            Selection.Paragraphs(1).Range.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.MoveRight Unit:=wdCharacter, Count:=Len(S1) + 1, Extend:=wdExtend
            Selection.Font.Bold = True
            Selection.Paragraphs(1).Range.Select
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        Else
            Selection.Paragraphs(1).Range.Select
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        End If
    Next para
    
    Selection.HomeKey Unit:=wdStory '插入点置于文档开始处
    Application.ScreenUpdating = True '恢复屏幕更新

End Sub

Sub A01_方括号内文本加粗()

    Application.ScreenUpdating = False '关闭屏幕更新
    C = Chr(93) '方括号]
    For Each para In ActiveDocument.Paragraphs
        TT = para.Range.Text
        L = Len(TT)
        P = InStr(1, TT, C, 1)
        If P > 0 Then
            S1 = Left(TT, P - 1)
            S2 = Right(TT, L - P)
            Selection.Paragraphs(1).Range.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.MoveRight Unit:=wdCharacter, Count:=Len(S1) + 1, Extend:=wdExtend
            Selection.Font.Bold = True
            Selection.Font.Color = wdColorBlue
            Selection.Paragraphs(1).Range.Select
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        Else
            Selection.Paragraphs(1).Range.Select
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        End If
    Next para
    
    Selection.HomeKey Unit:=wdStory '插入点置于文档开始处
    Application.ScreenUpdating = True '恢复屏幕更新

End Sub

Sub A01_冒号前文本加粗()

    Application.ScreenUpdating = False '关闭屏幕更新

    C = Chr(-23622) '冒号：
    For Each para In ActiveDocument.Paragraphs
        TT = para.Range.Text
        L = Len(TT)
        P = InStr(1, TT, C, 1)
        If P > 0 Then
            S1 = Left(TT, P - 1)
            S2 = Right(TT, L - P)
            Selection.Paragraphs(1).Range.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.MoveRight Unit:=wdCharacter, Count:=Len(S1) + 1, Extend:=wdExtend
            Selection.Font.Bold = True
            Selection.Font.Color = wdColorBlue
            Selection.Paragraphs(1).Range.Select
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        Else
            Selection.Paragraphs(1).Range.Select
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        End If
    Next para
    
    Selection.HomeKey Unit:=wdStory '插入点置于文档开始处
    Application.ScreenUpdating = True '恢复屏幕更新

End Sub


Sub 插入页码()
'
' 宏在 2009-2-3 由 戴宏国: 录制
'
    Selection.Sections(1).Footers(1).PageNumbers.Add Pagenumberalignment:= _
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
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:="—"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Font.Size = 12
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub

Sub A01_发文落款()  '宏命令名称
    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(14.07), Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
    Selection.TypeText Text:=Chr(9) & "统计资料管理中心" & Chr(13)
    Selection.InsertDateTime DateTimeFormat:=Chr(9) & "EEEE年O月A日", InsertAsField:=False
End Sub

Sub A01_处理表头单位()

'    Application.Run MacroName:="Normal.NewMacros.Macro5"
'    Application.Run MacroName:="A01_发文落款"
    DW = "单位"
    Selection.Tables(1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.Paragraphs(1).Range.Select
    T = Selection.Paragraphs(1).Range.Text
    T = Left(T, Len(T) - 1)
    S2 = LTrim(T)
    p2 = InStr(1, S2, DW, 1)
    If p2 > 0 Then
    Selection.Delete
'    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.InsertRowsAbove 1
    Selection.Cells.Merge
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.TypeText Text:=S2
    Selection.SelectRow
        With Selection.ParagraphFormat
            .LeftIndent = CentimetersToPoints(0.1)
            .RightIndent = CentimetersToPoints(0.5)
            .LineSpacingRule = wdLineSpaceAtLeast
            .LineSpacing = 12
            .Alignment = wdAlignParagraphRight
        End With
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.5)
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Else
        Selection.MoveRight Unit:=wdCharacter, Count:=2
    End If
        
End Sub

Sub A01_批量处理表头单位()

End Sub

'    Application.Run MacroName:="Normal.NewMacros.Macro5"
'    Application.Run MacroName:="A01_发文落款"
    TN = ActiveDocument.Tables.Count
    If TN > 0 Then
    For j = 1 To TN
        ActiveDocument.Tables(j).Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Application.Run MacroName:="A01_处理表头单位"
    Next j
    End If
Sub A01_缩进转换为空格()


For Each PA In ActiveDocument.Paragraphs
    PA.Range.Select
    T = Selection.Paragraphs(1).Range.Text
    T = Left(T, Len(T) - 1)
    If Selection.ParagraphFormat.FirstLineIndent = 24 Then
        Selection.ParagraphFormat.FirstLineIndent = 0
        t1 = Chr(-24159) & Chr(-24159) & T
        Selection.TypeText Text:=t1
    End If
Next PA


End Sub

Sub A01_英汉对照()

'Application.Run MacroName:="段落顺序颠倒"

    Dim PS() As Variant
    pn = ActiveDocument.Paragraphs.Count
    ReDim PS(pn - 1)

    For j = 1 To pn
    TT = ActiveDocument.Paragraphs(j).Range.Text
    TT = Left(TT, Len(TT) - 1)
    PS(j - 1) = TT
    Next j
    
    Selection.WholeStory
    Selection.Delete
    
    Set MyRange = ActiveDocument.Range(start:=0, End:=0)
    ActiveDocument.Tables.Add Range:=MyRange, NumRows:=pn, NumColumns:=2
    Set mytable = ActiveDocument.Tables(1)
    
    For i = 0 To UBound(PS)
    mytable.Cell(i + 1, 1).Select
    Selection.TypeText Text:=PS(i)
    i = i + 1
    Next i

    For i = 0 To UBound(PS)
    mytable.Cell(i + 1, 2).Select
    Selection.TypeText Text:=PS(i + 1)
    i = i + 1
    Next i

End Sub

Sub A01_英汉对照N()

'Application.Run MacroName:="A01_英汉对照"

    Dim PT(), PS(), PK() As Variant
    pn = ActiveDocument.Paragraphs.Count
    ReDim PT(pn - 1)
    ReDim PS(pn / 2 - 1)
    ReDim PK(pn / 2 - 1)
    
    For j = 1 To pn
    TT = ActiveDocument.Paragraphs(j).Range.Text
    TT = Left(TT, Len(TT) - 1)
    PT(j - 1) = TT
    Next j

    For i = 0 To (UBound(PT) - 1) / 2
    txt1 = PT(2 * i)
    txt2 = PT(2 * i + 1)
    PS(i) = txt1
    PK(i) = txt2
    Next i
    
'    For i = 0 To (UBound(PT) - 1) / 2
'    txt = PT(2 * i + 1)
'    PK(i) = txt
'    Next i
    
    Selection.WholeStory
    Selection.Delete
    Application.Run MacroName:="页面设置"
    
    Set MyRange = ActiveDocument.Range(start:=0, End:=0)
    ActiveDocument.Tables.Add Range:=MyRange, NumRows:=pn / 2, NumColumns:=2
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    ActiveDocument.Tables(1).Columns(1).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(1).PreferredWidth = 60
    ActiveDocument.Tables(1).Columns(2).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(2).PreferredWidth = 40
    
    Set mytable = ActiveDocument.Tables(1)

    For i = 0 To UBound(PS)
    mytable.Cell(i + 1, 1).Select
    Selection.TypeText Text:=PS(i)
    Next i

    For i = 0 To UBound(PK)
    mytable.Cell(i + 1, 2).Select
    Selection.TypeText Text:=PK(i)
    Next i


End Sub

Sub A01_表格转置N()

'Application.Run MacroName:="A01_英汉对照N"
    Dim RN, CN, i, j As Integer
    Dim TC(), TF() As String
    Dim MyRange As Range

    If ActiveDocument.Tables.Count >= 1 Then
        Set mytable = ActiveDocument.Tables(1)
        RN = mytable.Rows.Count
        CN = mytable.Columns.Count
        ReDim TC(RN, CN)
        ReDim TF(CN, RN)
        For i = 1 To RN
        For j = 1 To CN
        TT = mytable.Cell(i, j).Range.Text
        TT = Left(TT, Len(TT) - 2)
        TC(i, j) = TT
        TF(j, i) = TT
        Next j
        Next i
    End If
    
    Set MyRange = ActiveDocument.Range(start:=0, End:=0)
    ActiveDocument.Tables.Add Range:=MyRange, NumRows:=CN, NumColumns:=RN
    Set tbn = ActiveDocument.Tables(1)

    For i = 1 To CN
    For j = 1 To RN
    tbn.Cell(i, j).Range.Text = TF(i, j)
    Next j
    Next i

End Sub


Sub A01_表格到数组()
    Dim intCells As Integer
    Dim celTable As Cell
    Dim strCells() As String
    Dim intCount As Integer
    Dim rngText As Range

    If ActiveDocument.Tables.Count >= 1 Then
        With ActiveDocument.Tables(1).Range
            intCells = .Cells.Count
            ReDim strCells(intCells)
            intCount = 1
            For Each celTable In .Cells
                Set rngText = celTable.Range
                rngText.MoveEnd Unit:=wdCharacter, Count:=-1
                strCells(intCount) = rngText
                intCount = intCount + 1
            Next celTable
        End With
    End If
End Sub


Sub MyMacro1()
'
' Macro6 Macro
' 宏在 2009-4-1 由 便笺 录制
'
    Selection.InlineShapes.AddPicture FileName:= _
        "C:\Documents and Settings\ZLG005\My Documents\My Pictures\cp06.gif", LinkToFile:=False, SaveWithDocument:=True
    Selection.InlineShapes.AddPicture FileName:= _
        "C:\Documents and Settings\ZLG005\My Documents\My Pictures\mm016.jpg", LinkToFile:=False, SaveWithDocument:=True
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:="http://www.xinhuanet.com/", SubAddress:=""
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeParagraph
    Selection.TypeText Text:="国家统计局统计资料管理中心"
    Selection.TypeParagraph
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
'    ActiveDocument.Shapes.AddTextEffect(msoTextEffect8, "戴宏国", "宋体", 36#, msoFalse, msoFalse, 244.1, 247.1).Select
    Selection.TypeParagraph
End Sub


Sub 关闭样式自动更新()
'---------------------------------------------------
'功能：关闭活动文档所有样式的自动更新
'作者：戴宏国
'日期：2010年4月28日
'---------------------------------------------------
   
    Dim update As Style
    Set Updates = ActiveDocument.Styles
    For Each update In Updates
        If update.Type = wdStyleTypeParagraph Then
            update.AutomaticallyUpdate = False
        End If
    Next
End Sub

Sub ToggleInterpunction() '中英文标点互换
Dim ChineseInterpunction() As Variant, EnglishInterpunction() As Variant
Dim myArray1() As Variant, myArray2() As Variant, strFind As String, strRep As String
Dim msgResult As VbMsgBoxResult, N As Byte
'定义一个中文标点的数组对象
ChineseInterpunction = Array("、", "。", "，", "；", "：", "？", "！", "……", "-", "～", "（", "）", "《", "》")
'定义一个英文标点的数组对象
EnglishInterpunction = Array(",", ".", ",", ";", ":", "?", "!", "…", "-", "~", "(", ")", "&lt;", "&gt;")
'提示用户交互的MSGBOX对话框
msgResult = MsgBox("您想中英标点互换吗?按Y将中文标点转为英文标点,按N将英文标点转为中文标点!", vbYesNoCancel)
Select Case msgResult
Case vbCancel
Exit Sub '如果用户选择了取消按钮,则退出程序运行
Case vbYes '如果用户选择了YES,则将中文标点转换为英文标点
myArray1 = ChineseInterpunction
myArray2 = EnglishInterpunction
'strFind = " " ( * ) " "
strRep = """\1"""
Case vbNo '如果用户选择了NO,则将英文标点转换为中文标点
myArray1 = EnglishInterpunction
myArray2 = ChineseInterpunction
strFind = """(*)"""
'strRep = ""\1""
End Select
Application.ScreenUpdating = False '关闭屏幕更新
For N = 0 To UBound(ChineseInterpunction) '从数组的下标到上标间作一个循环
With ActiveDocument.Content.Find
.ClearFormatting '不限定查找格式
.MatchWildcards = False '不使用通配符
'查找相应的英文标点,替换为对应的中文标点
.Execute FindText:=myArray1(N), replacewith:=myArray2(N), Replace:=wdReplaceAll
End With
Next
With ActiveDocument.Content.Find
.ClearFormatting '不限定查找格式
.MatchWildcards = True '使用通配符
.Execute FindText:=strFind, replacewith:=strRep, Replace:=wdReplaceAll
End With
Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub 标点转换E_C()
    Selection.WholeStory
    Dim A As Variant, B As Variant
    
    x2 = Chr(-24157) '。
    x3 = Chr(-23636) '，
    x4 = Chr(-23621) '；
    x5 = Chr(-23622) '：
    x6 = Chr(-23617) '？
    x7 = Chr(-23647) '！
    x8 = Chr(-24150) '"—"
    x9 = Chr(-24149) '"～"
    x10 = Chr(-23640) '"（"
    x11 = Chr(-23639) '"（"
    A = Array(x2, x3, x4, x5, x6, x7, x8, x9, x10, x11)
    
    y2 = Chr(46) '.
    y3 = Chr(44) ',
    y4 = Chr(59) ';
    y5 = Chr(58) ':
    y6 = Chr(63) '?
    y7 = Chr(33) '!
    y8 = Chr(45) '-
    y9 = Chr(126) '~
    y10 = Chr(40)
    y11 = Chr(41)
    B = Array(y2, y3, y4, y5, y6, y7, y8, y9, y10, y11)
    
    For i = 0 To UBound(A)
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = B(i)
        .Replacement.Text = A(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
End Sub
Sub 标点转换C_E()
    Selection.WholeStory
    Dim A As Variant, B As Variant
    
    x1 = Chr(-24158) '、
    x2 = Chr(-24157) '。
    x3 = Chr(-23636) '，
    x4 = Chr(-23621) '；
    x5 = Chr(-23622) '：
    x6 = Chr(-23617) '？
    x7 = Chr(-23647) '！
    x8 = Chr(-24150) '"—"
    x9 = Chr(-24149) '"～"
    x10 = Chr(-23640) '"（"
    x11 = Chr(-23639) '"（"
    A = Array(x1, x2, x3, x4, x5, x6, x7, x8, x9, x10, x11)
    
    y1 = Chr(44) ',
    y2 = Chr(46) '.
    y3 = Chr(44) ',
    y4 = Chr(59) ';
    y5 = Chr(58) ':
    y6 = Chr(63) '?
    y7 = Chr(33) '!
    y8 = Chr(45) '-
    y9 = Chr(126) '~
    y10 = Chr(40)
    y11 = Chr(41)
    B = Array(y1, y2, y3, y4, y5, y6, y7, y8, y9, y10, y11)
    
    For i = 0 To UBound(A)
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = A(i)
        .Replacement.Text = B(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
End Sub

Sub A01_删字母()

    If Len(Selection.Range.Text) = 0 Then Selection.WholeStory
    For i = 65 To 90
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(i)
        .Replacement.Text = ""
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    
    For i = 97 To 122
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(i)
        .Replacement.Text = ""
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i

End Sub

Sub 扫描文本标准化()
    
    Dim MyRange As Range
        If Len(Selection.Range.Text) = 0 Then
            Selection.WholeStory
        End If
    Set MyRange = Selection.Range
    
    With MyRange
    
    回车替换
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(32) & Chr(32) & Chr(32) & Chr(32)
        .Replacement.Text = "^p"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    删除空格K
    段前加空
    
    End With

End Sub
Sub A01_两字人名()
    For Each par In ActiveDocument.Content.Paragraphs
      T = Trim(par.Range.Text)
            If Len(T) = 2 Then
                t1 = Left(par, 1)
                par.Range.Text = t1
            End If
    Next
End Sub

Sub A01_批量读取答题()
    
    Application.DisplayAlerts = wdAlertsNone

Set fs = Application.FileSearch
With fs
    .LookIn = "D:\开放日"
    .FileName = "*.doc"
    If .Execute(SortBy:=msoSortByFileName, _
    SortOrder:=msoSortOrderAscending) > 0 Then
        For i = 1 To .FoundFiles.Count
    Documents.Open FileName:=.FoundFiles(i)
    On Error Resume Next
    TXT = ActiveDocument.Paragraphs(1).Range.Text
    If Left(TXT, Len(TXT) - 1) = "中国统计开放日" Then
    
        A01_读取答题
        ActiveWindow.Close wdDoNotSaveChanges
'        Selection.WholeStory
'        Selection.EndKey
'        Selection.TypeParagraph
'        Selection.TypeText Text:=TT
    Else
        ActiveWindow.Close wdDoNotSaveChanges
    End If
    
        Next i
    Else
        MsgBox "没找到可用文档"
    End If
End With
    
End Sub

Sub A01_读取答题()
    

    A00_网页格式
    Selection.WholeStory
    删除空格K
    
    t1 = ActiveDocument.Paragraphs(3).Range.Text
    t2 = ActiveDocument.Paragraphs(8).Range.Text
    t3 = ActiveDocument.Paragraphs(13).Range.Text
    t4 = ActiveDocument.Paragraphs(19).Range.Text
    t5 = ActiveDocument.Paragraphs(28).Range.Text
    t6 = ActiveDocument.Paragraphs(29).Range.Text
    t7 = ActiveDocument.Paragraphs(30).Range.Text
    t8 = ActiveDocument.Paragraphs(31).Range.Text
    t9 = ActiveDocument.Paragraphs(32).Range.Text
    t10 = ActiveDocument.Paragraphs(35).Range.Text
    t11 = ActiveDocument.Paragraphs(18).Range.Text
    t12 = ActiveDocument.Paragraphs(25).Range.Text
    
    A1 = Trim(Right(Left(t1, 25), 1))
    A2 = Trim(Right(Left(t2, 37), 1))
    A3 = Trim(Right(Left(t3, Len(t3) - 2), Len(t3) - 22))
    A4 = Trim(Right(Left(t4, Len(t4) - 2), Len(t4) - 28))
    A5 = Trim(Right(Left(t5, Len(t5) - 1), Len(t5) - 8))
    A6 = Trim(Right(Left(t6, Len(t6) - 1), Len(t6) - 10))
    A7 = Trim(Right(Left(t7, Len(t7) - 1), Len(t7) - 8))
    A8 = Trim(Right(Left(t8, Len(t8) - 1), Len(t8) - 10))
    A9 = Trim(Right(Left(t9, Len(t9) - 2), Len(t9) - 12))
    A10 = Trim(Right(Left(t10, Len(t10) - 2), Len(t10) - 12))
    A11 = Trim(Right(Left(t11, Len(t11) - 1), Len(t11) - 12))
    A12 = Trim(Right(Left(t12, Len(t12) - 1), Len(t12) - 12))
    
    TT1 = "省(自治区、直辖市)"
    TT2 = "市(自治州、地区、盟)"
    TT3 = "县(区、特区、旗)"
    
    P1 = InStr(1, A8, TT1, 1)
    A81 = Left(A8, P1 - 1)
    A82 = Right(A8, Len(A8) - Len(TT1) - P1 + 1)
    p2 = InStr(1, A82, TT2, 1)
    A83 = Left(A82, p2 - 1)
    A84 = Right(A82, Len(A82) - Len(TT2) - p2 + 1)
    A85 = Left(A84, Len(A84) - 9)
    
    TT = A1 & Chr(9) & A2 & Chr(9) & A3 & Chr(9) & A4 & Chr(9) & A5 & Chr(9) & "S" & A6 & Chr(9) & A7 & Chr(9) & A81 & A83 & A85 & Chr(9) & A9 & Chr(9) & A10 & Chr(9) & A11 & Chr(9) & A12
    
    'MsgBox TT
    
    Set MyDoc1 = ActiveDocument
    
    For Each doc In Documents
        If InStr(1, doc.name, "答题汇总.doc", 1) Then
            doc.Activate
            docFound = True
            Exit For
        Else
            docFound = False
        End If
    Next doc
    If docFound = False Then Documents.Open FileName:="D:\00 F2011\答题汇总.doc"
    
    Selection.WholeStory
    Selection.EndKey
    Selection.TypeParagraph
    Selection.TypeText Text:=TT
    Set MyDoc2 = ActiveDocument
    MyDoc1.Activate
    
    

End Sub
Sub 页边距2厘米()
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
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

Sub name()
    For Each par In ActiveDocument.Content.Paragraphs
      T = Trim(par.Range.Text)
            If Len(T) = 2 Then
            t1 = lef(par, 1)
            par.Range.Text = t1
            Next
End Sub
Sub 资料中心()
'
' 资料中心 Macro
' 宏在 2011-10-17 由 戴宏国: 录制
'
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.TypeText Text:="资料中心"
    Selection.TypeParagraph
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DATE  \@ ""EEEE年O月A日"" ", PreserveFormatting:=True
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(0.74)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(8.89) _
        , Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
    Selection.HomeKey Unit:=wdLine
    Selection.TypeText Text:=vbTab
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.HomeKey Unit:=wdLine
    Selection.ParagraphFormat.TabStops.ClearAll
    ActiveDocument.DefaultTabStop = CentimetersToPoints(0.74)
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(8.89) _
        , Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
    Selection.TypeText Text:=vbTab
End Sub



Sub 右缩进半厘米()
    Selection.ParagraphFormat.RightIndent = CentimetersToPoints(0.5)
End Sub
Sub 右缩进1厘米()
    Selection.ParagraphFormat.RightIndent = CentimetersToPoints(1)
End Sub
Sub 右缩进1厘米半()
    Selection.ParagraphFormat.RightIndent = CentimetersToPoints(1.5)
End Sub
Sub 右缩进2厘米()
    Selection.ParagraphFormat.RightIndent = CentimetersToPoints(2)
End Sub
Sub Macro8()
'
' Macro8 Macro
' 宏在 2013-2-22 由 戴宏国:分发 录制
'
    CommandBars.Add(name:="我的表格工具").Visible = True
End Sub
Sub Macro13()
'
' Macro13 Macro
' 宏在 2013-2-22 由 戴宏国:分发 录制
'
    CommandBars.Add(name:="我的编辑工具").Visible = True
    CommandBars("我的编辑工具").Controls.Add Type:=msoControlButton, Before:=1
End Sub


Sub hypertest()
   
    Set MyRange = ActiveDocument.Range(start:=ActiveDocument.Content.End - 1)
    Count = 0
    For Each aHyperlink In ActiveDocument.Hyperlinks
        Count = Count + 1
        With MyRange
            .InsertAfter "Hyperlink #" & Count & vbTab
            .InsertAfter aHyperlink.Address
            .InsertParagraphAfter
        End With
    Next aHyperlink

End Sub

Sub hypertest_New()
   
    Dim MyRange As Range
    Dim i As Integer
    Dim PS() As Variant
    Dim WZ() As Variant
    
    Set MyRange = ActiveDocument.Range(start:=ActiveDocument.Content.End - 1)
        MyRange.InsertParagraphAfter
    
    i = 0
    ReDim Preserve PS(1)
    ReDim Preserve WZ(1)
    
    Application.ScreenUpdating = False '关闭屏幕更新
    
    For Each aHyperlink In ActiveDocument.Hyperlinks
        PS(i) = aHyperlink.Address
        WZ(i) = aHyperlink.TextToDisplay
        i = i + 1
    '    MsgBox "i " & i & " 链接 " & aHyperlink.Address
        ReDim Preserve PS(UBound(PS) + 1)
        ReDim Preserve WZ(UBound(WZ) + 1)
    Next aHyperlink
        'MsgBox UBound(PS)
    For j = 0 To UBound(PS) - 2
        With MyRange
            .InsertAfter "链接" & vbTab & PS(j) & vbTab & WZ(j) & Chr(13)
            '.InsertParagraphAfter
        End With
    Next j
    
    Application.ScreenUpdating = True '打开屏幕更新
    
End Sub

Sub 网页页面()
'
' 网页页面 Macro
' 宏在 2013-4-17 由 戴宏国: 录制
'
    With ActiveDocument.Styles(wdStyleNormal).Font
        If .NameFarEast = .NameAscii Then
            .NameAscii = ""
        End If
        .NameFarEast = ""
    End With
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientLandscape
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(4.85)
        .RightMargin = CentimetersToPoints(4.85)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
        .PageWidth = CentimetersToPoints(29.7)
        .PageHeight = CentimetersToPoints(21)
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

Sub B01_选定文字变为上标()
    Selection.Font.Superscript = True
    Selection.Font.Color = wdColorBlue
End Sub

Sub B01_方括号内文本变为上标()
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next '忽略错误
    
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory

    For i = 1 To 100
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="[" & i & "]"
        If .Found = True Then
            Selection.Font.Superscript = True
            Selection.Font.Color = wdColorBlue
        Else
            Exit For
        End If
    End With
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    Next i
    
    Application.ScreenUpdating = True '打开屏幕更新

End Sub

Sub B01_行首缩进转变纯中文空格()

    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next '忽略错误

    Set MyRange = Selection.Range
        NRow = MyRange.Rows.Count
        MyRange.Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
    For N = 1 To NRow
        Selection.Cells(1).Range.Select
        KN = Selection.ParagraphFormat.CharacterUnitFirstLineIndent
    If KN > 0 Then
        KG = ""
        For i = 1 To Int(KN)
            KG = KG & "　"
        Next i
    End If
    Selection.Cells(1).Range.Select
    Ctxt = Selection.Range.Text
    TT = Left(Ctxt, Len(Ctxt) - 2)
    Selection.Range.Text = KG & TT
    Selection.Cells(1).Range.Select
    Selection.MoveDown Unit:=wdLine, Count:=1
    KN = 0
    KG = ""
    Next N
    
    MyRange.Select
    
    A01_取消行首缩进
    A01_取消行首缩进
    
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True '打开屏幕更新


End Sub

Sub F01_翻译新闻稿附注()
'
' 定义三个变量：NM -- 文件名； N -- 图片数  T--表格数
    Dim NM As String
        
    ' 取得当前文件名
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    
    Application.ScreenUpdating = False '关闭屏幕更新

    Documents.Open FileName:="D:\00 Word_Dot\附注.doc"
    
    Set MyDoc1 = Application.ActiveWindow.Document

    Application.ScreenUpdating = False '关闭屏幕更新

    H = ActiveDocument.Tables(1).Rows.Count
    For i = 1 To H
    C = ActiveDocument.Tables(1).Columns(1).Cells(i).Range.Text
    C = Left(C, Len(C) - 2)
'    e = ActiveDocument.Tables(1).Columns(2).Cells(i).Range.Text
    ActiveDocument.Tables(1).Columns(2).Cells(i).Range.Copy

    MyDOC.Activate
    
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=C
        If .Found = True Then
        Selection.Paste
        End If
    End With

    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    
    MyDoc1.Activate
    
    Next i
    
    MyDoc1.Close
    MyDOC.Activate
    
Application.ScreenUpdating = True '关闭屏幕更新

End Sub
Sub A01_加超链接()

    Dim ML, FN As String
    Set mytable = ActiveDocument.Tables(1)
    R = mytable.Rows.Count
    C = mytable.Columns.Count
    
    Application.ScreenUpdating = False '关闭屏幕更新
    
    ML = ""
    j = 1
    For j = 1 To R
    pn = mytable.Rows(j).Cells(2).Range.Text
    FN1 = Left(pn, Len(pn) - 2)
    mytable.Rows(j).Cells(1).Select
    td = Selection.Range.Text
    TD1 = Left(td, Len(td) - 2)
    Selection.Delete
    
    Selection.Hyperlinks.Add Anchor:=Selection.Range, Address:=ML & FN1, TextToDisplay:=TD1
   
Next j

    Application.ScreenUpdating = True '打开屏幕更新


End Sub


Sub 英文引号转换为中文引号()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    With Selection.Find
        .Text = """"
        .Forward = True
        .Wrap = wdStop
        .MatchByte = True
    End With
    
    With Selection
        While .Find.Execute
            .Text = ChrW(8220)
            .Find.Execute
            .Text = ChrW(8221)
            Wend
    End With
End Sub

Sub A01_千分位()
'本代码旨在解决WORD中数据转化为千分位
'数据限定要求:-922,337,203,685,477.5808 到 922,337,203,685,477.5807
'转化结果1000以上数据以千分位计算,小数点右侧保留二位小数;1000以下数据不变
Dim MyRange As Range, i As Byte, myValue As Currency
On Error Resume Next '忽略错误
st = VBA.Timer '计时器
Application.ScreenUpdating = False '关闭屏幕更新

NextFind: Set MyRange = ActiveDocument.Content '定义为主文档文字部分
With MyRange.Find '查找
    .ClearFormatting '清除格式
    .Text = "[0-9]{4,15}" '4到15位数据
    .MatchWildcards = True '使用通配符
Do While .Execute '每次查找成功
    i = 2 '起始值为2
    '如果是有小数点
    If MyRange.Next(wdCharacter, 1) = "." Then
    '进行一个未知循环
        While MyRange.Next(wdCharacter, i) Like "#"
            i = i + 1 '只要是[0-9]任意数字则累加
        Wend
        '重新定义RANGE对象
        MyRange.SetRange MyRange.start, MyRange.End + i - 1
    End If
    myValue = VBA.Val(MyRange) '保险起见转换为数据,也可省略
    MyRange = VBA.Format(myValue, "Standard") '转为千分位格式
    GoTo NextFind '转到指定行
    Loop
    End With

Application.ScreenUpdating = True '恢复屏幕更新
MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒" '显示程序运行所花费的时间

End Sub

Sub A01_设置图片大小比例() '设置图片大小为当前的百分比
Dim N '图片个数
Dim picwidth
Dim picheight
If Selection.Type = wdSelectionNormal Then
On Error Resume Next '忽略错误
For N = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes类型图片
picheight = ActiveDocument.InlineShapes(N).Height
picwidth = ActiveDocument.InlineShapes(N).Width
ActiveDocument.InlineShapes(N).Height = picheight * 0.5 '设置高度
ActiveDocument.InlineShapes(N).Width = picwidth * 0.5 '设置宽度
Next N
For N = 1 To ActiveDocument.Shapes.Count 'Shapes类型图片
picheight = ActiveDocument.Shapes(N).Height
picwidth = ActiveDocument.Shapes(N).Width
ActiveDocument.Shapes(N).Height = picheight * 0.5 '设置高度倍数
ActiveDocument.Shapes(N).Width = picwidth * 0.5 '设置宽度倍数
Next N

Else: End If
End Sub


Sub A01_设置图片大小值() '设置图片大小为固定值
Dim N '图片个数
On Error Resume Next '忽略错误
For N = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes类型图片
ActiveDocument.InlineShapes(N).Height = 200 '设置图片高度为 400px
ActiveDocument.InlineShapes(N).Width = 300 '设置图片宽度 300px
Next N
For N = 1 To ActiveDocument.Shapes.Count 'Shapes类型图片
ActiveDocument.Shapes(N).Height = 200 '设置图片高度为 400px
ActiveDocument.Shapes(N).Width = 300 '设置图片宽度 300px
Next N
End Sub

Sub A01_图片版式转换()
'* ＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋＋
'* Created By SHOUROU@ExcelHome 2007-12-11 5:28:26
'仅测试于System: Windows NT Word: 11.0 Language: 2052
'№ 0281^The Code CopyIn [ThisDocument-ThisDocument]^'
'* －－－－－－－－－－－－－－－－－－－－－－－－－－－－－
'Option Explicit Dim oShape As Variant, shapeType As WdWrapType
On Error Resume Next
If MsgBox("Y将图片由嵌入式转为浮动式,N将图片由浮动式转为嵌入式", 68) = 6 Then
shapeType = Val(InputBox(Prompt:="请输入图片版式:0=四周型,1=紧密型, " & vbLf & _
"3=衬于文字下方,4=浮于文字上方", Default:=0))
For Each oShape In ActiveDocument.InlineShapes
Set oShape = oShape.ConvertToShape
With oShape
Select Case shapeType
Case 0, 1
.WrapFormat.Type = shapeType
Case 3
.WrapFormat.Type = 3
.ZOrder 5
Case 4
.WrapFormat.Type = 3
.ZOrder 4
Case Else
Exit Sub
End Select
.WrapFormat.AllowOverlap = False '不允许重叠
End With
Next
Else
For Each oShape In ActiveDocument.Shapes
oShape.ConvertToInlineShape
Next
End If
End Sub

Sub A01_GetChineseNum2()
'把数字转化为汉字大写人民币
Dim Numeric As Currency, IntPart As Long, DecimalPart As Byte, MyField As Field, Label As String
Dim Jiao As Byte, Fen As Byte, Oddment As String, Odd As String, MyChinese As String
Dim strNumber As String
Const ZWDX As String = "壹贰叁肆伍陆柒捌玖零" '定义一个中文大写汉字常量
On Error Resume Next '错误忽略
If Selection.Type = wdSelectionNormal Then

With Selection
strNumber = VBA.Replace(.Text, " ", "")
Numeric = VBA.Round(VBA.CCur(strNumber), 2) '四舍五入保留小数点后两位
'判断是否在表格中
If .Information(wdWithInTable) Then _
.MoveRight Unit:=wdCell Else .MoveRight Unit:=wdCharacter
'对数据进行判断,是否在指定的范围内
If VBA.Abs(Numeric) > 2147483647 Then MsgBox "数值超过范围!", _
vbOKOnly + vbExclamation, "Warning": Exit Sub
IntPart = Int(VBA.Abs(Numeric)) '定义一个正整数
Odd = VBA.IIf(IntPart = 0, "", "圆") '定义一个STRING变量
'插入中文大写前的标签
Label = VBA.IIf(Numeric = VBA.Abs(Numeric), "人民币金额大写：", "人民币金额大写：负")
'对小数点后面二位数进行择定
DecimalPart = (VBA.Abs(Numeric) - IntPart) * 100
Select Case DecimalPart
Case Is = 0 '如果是0,即是选定的数据为整数
Oddment = VBA.IIf(Odd = "", "", Odd & "整")
Case Is < 10 '<10,即是零头是分
Oddment = VBA.IIf(Odd <> "", "圆零" & VBA.Mid(ZWDX, DecimalPart, 1) & "分", _
VBA.Mid(ZWDX, DecimalPart, 1) & "分")
Case 10, 20, 30, 40, 50, 60, 70, 80, 90 '如果是角整
Oddment = "圆" & VBA.Mid(ZWDX, DecimalPart / 10, 1) & "角整"
Case Else '既有角,又有分的情况
Jiao = VBA.Left(CStr(DecimalPart), 1) '取得角面值
Fen = VBA.Right(CStr(DecimalPart), 1) '取得分面值
Oddment = Odd & VBA.Mid(ZWDX, Jiao, 1) & "角" '转换为角的中文大写
Oddment = Oddment & VBA.Mid(ZWDX, Fen, 1) & "分" '转换为分的中文大写
End Select
'指定区域插入中文大写格式的域
Set MyField = .Fields.Add(Range:=.Range, Text:="= " & IntPart & " \*CHINESENUM2")
MyField.Select '选定域(最后是用指定文本覆盖选定区域)
'如果仅有角分情况下,Mychinese为""
MyChinese = VBA.IIf(MyField.Result <> "零", MyField.Result, "")
.Text = Label & MyChinese & Oddment
End With
Else
MsgBox "您没有选择数字，请选择数字！"
End If
End Sub

Sub A01_中英文标点互换() '中英文标点互换
Dim ChineseInterpunction() As Variant, EnglishInterpunction() As Variant
Dim myArray1() As Variant, myArray2() As Variant, strFind As String, strRep As String
Dim msgResult As VbMsgBoxResult, N As Byte
'定义一个中文标点的数组对象
ChineseInterpunction = Array("、", "。", "，", "；", "：", "？", "！", "……", "—", "～", "（", "）", "《", "》")
'定义一个英文标点的数组对象
EnglishInterpunction = Array(",", ".", ",", ";", ":", "?", "!", "…", "-", "~", "(", ")", "&lt;", "&gt;")
'提示用户交互的MSGBOX对话框
msgResult = MsgBox("您想中英标点互换吗?按Y将中文标点转为英文标点,按N将英文标点转为中文标点!", vbYesNoCancel)
Select Case msgResult
Case vbCancel
Exit Sub '如果用户选择了取消按钮,则退出程序运行
Case vbYes '如果用户选择了YES,则将中文标点转换为英文标点
myArray1 = ChineseInterpunction
myArray2 = EnglishInterpunction
strFind = "“(*)”"
strRep = """\1"""
Case vbNo '如果用户选择了NO,则将英文标点转换为中文标点
myArray1 = EnglishInterpunction
myArray2 = ChineseInterpunction
strFind = """(*)"""
strRep = "“\1”"
End Select
Application.ScreenUpdating = False '关闭屏幕更新
For N = 0 To UBound(ChineseInterpunction) '从数组的下标到上标间作一个循环
With ActiveDocument.Content.Find
.ClearFormatting '不限定查找格式
.MatchWildcards = False '不使用通配符
'查找相应的英文标点,替换为对应的中文标点
.Execute FindText:=myArray1(N), replacewith:=myArray2(N), Replace:=wdReplaceAll
End With
Next
With ActiveDocument.Content.Find
.ClearFormatting '不限定查找格式
.MatchWildcards = True '使用通配符
.Execute FindText:=strFind, replacewith:=strRep, Replace:=wdReplaceAll
End With
Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub 设置图片大小为原始大小()
Dim N '图片个数
Dim picwidth
Dim picheight
On Error Resume Next '忽略错误
For N = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes类型图片
ActiveDocument.InlineShapes(N).Reset
Next N
For N = 1 To ActiveDocument.Shapes.Count 'Shapes类型图片
ActiveDocument.Shapes(N).Select
Selection.ShapeRange.ScaleHeight 1#, msoTrue, msoScaleFromTopClientHeigh
Selection.ShapeRange.ScaleWidth 1#, msoTrue, msoScaleFromTopClientwidth
Next N
End Sub

Sub A01_mySaveAs()
'

Dim i As Long, st As Single, mypath As String, fs As FileSearch
Dim MyDOC As Document, N As Integer
Dim strpara1 As String, strpara2 As String, docname As String, A

On Error GoTo hd
With Application.FileDialog(msoFileDialogFilePicker)
.Title = "请选定任一文件，确定后将重命名全部WORD文档"
If .Show <> -1 Then Exit Sub
st = Timer
mypath = .InitialFileName
End With

Application.ScreenUpdating = False
If Dir(mypath & "另存为", vbDirectory) = "" Then MkDir mypath & "另存为" '另存为文档的保存位置
Set fs = Application.FileSearch
With fs
.NewSearch
.LookIn = mypath
.FileType = msoFileTypeWordDocuments
If .Execute(msoSortByFileName) > 0 Then
For i = 1 To .FoundFiles.Count
If InStr(fs.FoundFiles(i), "~$") = 0 Then
Set MyDOC = Documents.Open(.FoundFiles(i), Visible:=False)
With MyDOC
strpara1 = Replace(.Paragraphs(1).Range.Text, Chr(13), "")
strpara1 = Left(strpara1, 10)
strpara2 = Replace(.Paragraphs(2).Range.Text, Chr(13), "")
If Len(strpara1) < 2 Or Len(strpara2) < 2 Then GoTo hd
docname = strpara1 & "_" & strpara2
docname = CleanString(docname)
For Each A In Array("\", "/", ":", "*", "?", """ ", "<", " >", "|")
docname = Replace(docname, A, "")
Next
.SaveAs mypath & "另存为\" & docname & ".doc"
N = N + 1
.Close
End With
End If
Next
End If
End With
MsgBox "共处理了" & fs.FoundFiles.Count & "个文档，保存于目标文件夹的名称为“另存为”的下一级文件夹中。" _
& vbCrLf & "处理时间：" & Format(Timer - st, "0") & "秒。"
Application.ScreenUpdating = True
Exit Sub

hd:
MsgBox "运行出现意外，程序终止！" & vbCrLf & "已处理文档数：" & N _
& vbCrLf & "出错文档：" & vbCrLf & fs.FoundFiles(i)
If Not MyDOC Is Nothing Then MyDOC.Close
End Sub


Sub A01_分式()
'
' 分式 Macro
' 设置选定分数,快捷键为"Alt+F"
'
Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
If Selection.Type = wdSelectionNormal Then
'Selection.Font.Italic = True
Selection.Cut
Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
PreserveFormatting:=False
Selection.MoveRight Unit:=wdCharacter, Count:=1
Selection.TypeBackspace
Selection.TypeText Text:="eq \f()"
Selection.MoveLeft Unit:=wdCharacter, Count:=1
Selection.Paste
'Selection.TypeText Text:=")"
Selection.Fields.update
Selection.MoveRight Unit:=wdCharacter, Count:=1
Else
MsgBox "您没有选择文字。"
End If
'
End Sub

Sub A01_弧()
'
' 弧 Macro
' 设置选定的两个字母上加弧
Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
If Selection.Type = wdSelectionNormal Then
Selection.Font.Italic = True
Selection.Cut
Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
PreserveFormatting:=False
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.TypeText Text:="eq \o(\s\up5(⌒"
Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
Selection.Font.Scaling = 150
Selection.MoveRight Unit:=wdCharacter, Count:=1
Selection.Font.Scaling = 100
Selection.TypeText Text:="),\s\do0("
Selection.Paste
Selection.TypeText Text:="))"
Selection.Fields.update
Selection.MoveRight Unit:=wdCharacter, Count:=1
Else
MsgBox "您没有选择文字。"
End If
'
End Sub

Sub A01_Password()
'
' 文件自动添加密码。
'
If ActiveDocument.WriteReserved = False Then
If MsgBox("是否为本文档添加密码？", vbYesNo) = vbYes Then With ActiveDocument
.Password = "123456"
.WritePassword = "123456"
End With

Else
End If
Else
End If
End Sub

Sub A01_Example()

    '根据文档字符数中重复频率排序字符并计数
    Dim iCount As Long, i As Long, N As Long
    Dim ochar As String, TempA As Variant, st As Single
    Dim Array_Keys() As Variant, Array_Items() As Variant

    st = VBA.Timer
    Set myDictionary = CreateObject("Scripting.Dictionary")
    MyString = ActiveDocument.Content.Text
    N = Len(MyString) - 1
    For i = 1 To N
        ochar = VBA.Mid(MyString, i, 1)
        If myDictionary.Exists(ochar) = False Then
            myDictionary.Add ochar, 1
        Else
            myDictionary(ochar) = myDictionary(ochar) + 1
        End If
    Next
    
    MyString = ""
    iCount = myDictionary.Count - 1
    Array_Keys = myDictionary.keys
    Array_Items = myDictionary.Items
    Set myDictionary = Nothing
    
    For i = 0 To iCount - 1
        For N = i + 1 To iCount
            If Array_Items(i) < Array_Items(N) Then
                TempA = Array_Items(N)
                Array_Items(N) = Array_Items(i)
                Array_Items(i) = TempA
                TempA = Array_Keys(N)
                Array_Keys(N) = Array_Keys(i)
                Array_Keys(i) = TempA
            End If
        Next N
    Next i
    For i = 0 To iCount
        MyString = MyString & Array_Keys(i) & Chr(9) & Array_Items(i) & Chr(13)
    Next
    
    ActiveDocument.Content.Text = MyString
    MsgBox "共有" & iCount & "个不重复的字符,用时" & VBA.Format(Timer - st, "0.00") & "秒"

End Sub

Sub A01_Test()

Dim bw, sw, i As Integer
Dim MyCell As Cell
Dim TXT, MyTXT As Variant
    st = VBA.Timer
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next '忽略错误
    
    ActiveDocument.Tables(1).Cell(1, 1).Select
    Selection.SelectColumn
    i = 1

    For Each MyCell In Selection.Cells
        TXT = MyCell.Range.Text
        bw = Left(TXT, 1)
        sw = Mid(TXT, 2, 1)
        If (bw = 0 Or bw = 4 Or bw = 7) And (sw = 0 Or sw = 4 Or sw = 7) Then
            MyTXT = "000"
        Else
            If (bw = 0 Or bw = 4 Or bw = 7) And (sw = 1 Or sw = 5 Or sw = 8) Then
                MyTXT = "100"
            Else
                If (bw = 0 Or bw = 4 Or bw = 7) And (sw = 2 Or sw = 6 Or sw = 9) Then
                    MyTXT = "200"
                Else
                    If (bw = 1 Or bw = 5 Or bw = 8) And (sw = 0 Or sw = 4 Or sw = 7) Then
                        MyTXT = "300"
                    Else
                        If (bw = 1 Or bw = 5 Or bw = 8) And (sw = 1 Or sw = 5 Or sw = 8) Then
                            MyTXT = "400"
                        Else
                            If (bw = 1 Or bw = 5 Or bw = 8) And (sw = 2 Or sw = 6 Or sw = 9) Then
                                MyTXT = "500"
                            Else
                                If (bw = 2 Or bw = 6 Or bw = 9) And (sw = 0 Or sw = 4 Or sw = 7) Then
                                    MyTXT = "600"
                                Else
                                    If (bw = 2 Or bw = 6 Or bw = 9) And (sw = 1 Or sw = 5 Or sw = 8) Then
                                        MyTXT = "700"
                                    Else
                                        If (bw = 2 Or bw = 6 Or bw = 9) And (sw = 2 Or sw = 6 Or sw = 9) Then
                                            MyTXT = "800"
                                        Else
                                            If bw = 3 Or sw = 3 Then
                                                MyTXT = "900"
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        ActiveDocument.Tables(1).Cell(i, 2).Range.Text = MyTXT
        i = i + 1
    Next MyCell
    
    Application.ScreenUpdating = True '恢复屏幕更新
    
    MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒"


End Sub

Sub A01_Test1()

Dim bw, sw, i As Integer
Dim MyCell As Cell
Dim TXT, MyTXT As Variant
    st = VBA.Timer
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next '忽略错误
    
    ActiveDocument.Tables(1).Cell(1, 1).Select
    Selection.SelectColumn
    i = 1

    For Each MyCell In Selection.Cells
        TXT = MyCell.Range.Text
        bw = Left(TXT, 1)
        sw = Mid(TXT, 2, 1)
        Select Case bw
            Case 0, 4, 7
                If sw = 0 Or sw = 4 Or sw = 7 Then MyTXT = "000"
                If sw = 1 Or sw = 5 Or sw = 8 Then MyTXT = "100"
                If sw = 2 Or sw = 6 Or sw = 9 Then MyTXT = "200"
                If sw = 3 Then MyTXT = "900"
            Case 1, 5, 8
                If sw = 0 Or sw = 4 Or sw = 7 Then MyTXT = "300"
                If sw = 1 Or sw = 5 Or sw = 8 Then MyTXT = "400"
                If sw = 2 Or sw = 6 Or sw = 9 Then MyTXT = "500"
                If sw = 3 Then MyTXT = "900"
            Case 2, 6, 9
                If sw = 0 Or sw = 4 Or sw = 7 Then MyTXT = "600"
                If sw = 1 Or sw = 5 Or sw = 8 Then MyTXT = "700"
                If sw = 2 Or sw = 6 Or sw = 9 Then MyTXT = "800"
                If sw = 3 Then MyTXT = "900"
            Case 3
                MyTXT = "900"
        End Select
        
        ActiveDocument.Tables(1).Cell(i, 2).Range.Text = MyTXT
        i = i + 1
    Next MyCell
    
    Application.ScreenUpdating = True '恢复屏幕更新
    
    MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒"

End Sub
Sub A01_Test2()

Dim bw, sw, i As Integer
Dim MyCell As Cell
Dim TXT, MyTXT As Variant
    st = VBA.Timer
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next '忽略错误
    
    ActiveDocument.Tables(1).Cell(1, 1).Select
    Selection.SelectColumn
    i = 1

    For Each MyCell In Selection.Cells
        TXT = MyCell.Range.Text
        bw = Left(TXT, 1)
        sw = Mid(TXT, 2, 1)
        Select Case bw
            Case 0, 4, 7
                Select Case sw
                    Case 0, 4, 7
                        MyTXT = "000"
                    Case 1, 5, 8
                        MyTXT = "100"
                    Case 2, 6, 9
                        MyTXT = "200"
                    Case 3
                        MyTXT = "900"
                End Select
            Case 1, 5, 8
                Select Case sw
                    Case 0, 4, 7
                        MyTXT = "300"
                    Case 1, 5, 8
                        MyTXT = "400"
                    Case 2, 6, 9
                        MyTXT = "500"
                    Case 3
                        MyTXT = "900"
                End Select
            Case 2, 6, 9
                Select Case sw
                    Case 0, 4, 7
                        MyTXT = "600"
                    Case 1, 5, 8
                        MyTXT = "700"
                    Case 2, 6, 9
                        MyTXT = "800"
                    Case 3
                        MyTXT = "900"
                End Select
            Case 3
                MyTXT = "900"
        End Select
        
        ActiveDocument.Tables(1).Cell(i, 2).Range.Text = MyTXT
        i = i + 1
    Next MyCell
    
    Application.ScreenUpdating = True '恢复屏幕更新
    
    MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒"

End Sub

Sub A01_地址分列()  '功能：将地址按街、路、号分列
    
    Application.ScreenUpdating = False '恢复屏幕更新
    On Error Resume Next

    Dim MyRange As Range
        If Len(Selection.Range.Text) = 0 Then Selection.WholeStory
    Set MyRange = Selection.Range
    street = "街"
    road = "路"
    txt2 = "号"
    TT = ""
    st = VBA.Timer
    
    With MyRange
        pn = MyRange.Paragraphs.Count
        For i = 1 To pn
            T = Trim(MyRange.Paragraphs(i).Range.Text)
            T = Left(T, Len(T) - 1)
            '查找街或路
            p2 = InStr(1, T, street, 1)
            If p2 = 0 Then p2 = InStr(1, T, road, 1)
            
            FK1 = Left(T, p2)
            'FK2 = Right(t, Len(t) - p2)
            FK2 = Mid(T, p2 + 1)
            '查"号"
            p3 = InStr(1, FK2, txt2, 1)
            FK3 = Left(FK2, p3)
            FK4 = Mid(FK2, p3 + 1)
            'FK4 = Right(FK2, Len(FK2) - p3)
            TT = TT & FK1 & Chr(9) & FK3 & Chr(9) & FK4 & Chr(13)
        Next i
    End With
    Application.ScreenUpdating = True '恢复屏幕更新
    ActiveDocument.Content.Text = TT
    MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒"

End Sub

Sub A01_批量修改一组指定文件夹下的所有文件()  '功能：对验收资料按统一要求进行标准化：标题4黑、正文5宋

    '设置程序运行环境
    Application.ScreenUpdating = False '关闭屏幕更新，提高程序运行速度
    On Error Resume Next '遇到错误忽略之
    st = VBA.Timer '计数器
    
    '第一步：把事先准备好的验收资料目录下的全部子目录信息读取到数组变量中
    Dim ML() As Variant '定义名为ML的数组变量
    Dim i, j As Integer '定义2个计数变量
    Documents.Open FileName:="F:\AB\目录信息2.doc"   '打开文件目录信息.doc
    i = 1 '初始化计数变量i
    For Each PA In ActiveDocument.Paragraphs '进行一个循环操作，遍历当前文档的每个段落
        TT = Left(Trim(PA.Range.Text), Len(Trim(PA.Range.Text)) - 1) '读取每行文本，不包括换行符
        ReDim Preserve ML(i) '重新定义数组变量维度，保留原数组内容
        ML(i - 1) = TT  '把读取的文本信息赋值给数组元素
        i = i + 1  '计数器加1
    Next PA '跳转到下一个段落，进行重复操作
    Documents.Close SaveChanges:=wdDoNotSaveChanges '关闭文档
    
    '第二步：对每个子目录下的文件进行操作，逐一打开文件，调整格式
    For j = 0 To UBound(ML) '循环操作，遍历数组中的每个元素
        
            ChangeFileOpenDirectory "F:\B01_DOC"
            Set fs = Application.FileSearch
            With fs
                .LookIn = ML(j)
                .FileName = "*.DOCX"
                If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) > 0 Then
                    For i = 1 To .FoundFiles.Count
                        FN = .FoundFiles(i)
                        Documents.Open (FN), Visible:=False
                        Documents(FN).Activate
                        A01_格式调整
                        Set MyDOC = Application.ActiveWindow.Document
                        NM = Left(MyDOC, Len(MyDOC) - 5)
                        PRE = "B"
                        NM1 = PRE & NM
                        ActiveDocument.SaveAs FileName:=NM1, FileFormat:=wdFormatDocument
                        Documents.Close SaveChanges:=wdDoNotSaveChanges
                    Next i
                End If
            End With
    Next j
    Application.ScreenUpdating = True '恢复屏幕更新
    Application.Visible = True '恢复文档可视
    ChangeFileOpenDirectory "D:\00 F2013"
    
    MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒" '显示程序运行所花费的时间

End Sub


Sub A01_格式调整()

    Selection.WholeStory '选中全部内容
        Selection.Font.name = "宋体"
        Selection.Font.Size = 9
        Selection.Font.name = "Times New Roman"
        Selection.Font.Color = wdColorBlack
        Selection.WholeStory '选中全部内容
        Selection.HomeKey Unit:=wdStory
    Selection.Paragraphs(1).Range.Select '选中标题行
        Selection.Font.name = "黑体"
        Selection.Font.Size = 14
        Selection.Font.Bold = True
        Selection.Font.Color = wdColorBlack

End Sub

Sub A01_修改指定目录下文件()  '功能：调试宏命令

    'A01_批量修改一组指定文件夹下的所有文件
        '设置程序运行环境
    Application.ScreenUpdating = False '关闭屏幕更新，提高程序运行速度
    On Error Resume Next '遇到错误忽略之
    st = VBA.Timer '计数器

            Set fs = Application.FileSearch
            With fs
                .LookIn = "F:\验收资料\验收申请表\61\"
                .FileName = "*.DOC"
                If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) > 0 Then
                    For i = 1 To .FoundFiles.Count
                        FN = .FoundFiles(i)
                        Documents.Open (FN), Visible:=False
                        Documents(FN).Activate
                            A01_格式调整
                            Documents(FN).Close SaveChanges:=wdSaveChanges

                    Next i
                End If
            End With
    Application.ScreenUpdating = True '恢复屏幕更新
    Application.Visible = True '恢复文档可视
    
    MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒" '显示程序运行所花费的时间

End Sub
Sub A01_测试02()  '功能：调试宏命令

    Application.ScreenUpdating = False '关闭屏幕更新，提高程序运行速度
    On Error Resume Next '遇到错误忽略之
    st = VBA.Timer '计数器
            
            ChangeFileOpenDirectory "F:\B01_DOC\"
            Set fs = Application.FileSearch
            With fs
                .LookIn = "F:\B01_DOC\"
                .FileName = "*.DOC"
                If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) > 0 Then
                    For i = 1 To .FoundFiles.Count
                        FN = .FoundFiles(i)
                        Documents.Open (FN), Visible:=False
                        Documents(FN).Activate
                        A01_格式调整
                        ActiveDocument.Save
                        Documents.Close SaveChanges:=wdSaveChanges
                    Next i
                End If
            End With
    Application.ScreenUpdating = True '恢复屏幕更新
    Application.Visible = True '恢复文档可视
    ChangeFileOpenDirectory "D:\00 F2013\"  '恢复打开文件的缺省路径
    
    MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒" '显示程序运行所花费的时间

End Sub

Sub TT_Ontime()

'本示例设置 15 秒后运行 my_Procedure 过程，从现在开始计时。

    Application.OnTime Now + TimeValue("00:00:15"), "my_Procedure"

End Sub

Sub A01_编单表()

   '程序编写：Frank
   '编写日期：2014年3月20日
   '程序功能：按指定要求对特定表格（11列足彩表）进行编辑。指定要求：
   '1.删除表格的第7、8、9、10列；
   '2.在表格最下面插入1空白行；
   '3.表格属性里单元格，行，列，表格里全部数据都是0，√去掉
   '4.表格内段落，段前段后行距为固定值10磅
   
   Application.ScreenUpdating = False '关闭屏幕更新, 以提高代码运行速度
   'st = VBA.Timer '计时器
   Dim MyTab As Table '定义一个名为MyTab的表格变量
   
   '第一步：判断插入点（即光标位置）是否处在表格内：
   '如果是，则执行程序代码；
   '如果否，则弹出提示窗口，提示用户将插入点置于表格内（任意单元格内）
   
   If Selection.Information(wdWithInTable) = True Then
        Set MyTab = Selection.Tables(1)
        CN = MyTab.Columns.Count '识别表格的列数
        RN = MyTab.Rows.Count '识别表格的行数
        If CN <> 11 Then '如果不是11列的表格，则弹出提示窗口，提示用户，对表格不做任何处理，退出宏命令
            MsgBox "本表格不符合指定条件（11列足彩表），不能执行本宏命令！请选择符合条件的表格，谢谢！"
        Else
            MyTab.Rows(RN).Select '选中最后一行
            Selection.InsertRowsBelow 1 '插入1行
            MyTab.Columns(7).Select
            Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
            Selection.Cut '删除表格的第7、8、9、10列
            MyTab.Select
            With Selection.ParagraphFormat '表格内段落，段前段后空为0，行距为固定值10磅
                .SpaceBefore = 0
                .SpaceBeforeAuto = False
                .SpaceAfter = 0
                .SpaceAfterAuto = False
                .LineSpacingRule = wdLineSpaceExactly
                .LineSpacing = 10
                .LineUnitBefore = 0
                .LineUnitAfter = 0
                .WordWrap = True
            End With
            '设置表格宽度为按内容自动调整
            MyTab.AutoFitBehavior (wdAutoFitContent)
            MyTab.AutoFitBehavior (wdAutoFitContent)
            '把光标定位在最后一行的第2个单元格中
            MyTab.Rows(RN + 1).Select '选中最后一行
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        End If
   Else
        MsgBox "【注意】插入点（光标）不在表格中！" & Chr(13) & _
           "　　　　请将插入点（光标）置于表格中任意单元格中，" & Chr(13) & _
           "　　　　然后再执行本宏命令，谢谢！"
   End If
   
      Application.ScreenUpdating = True '恢复屏幕更新
      'MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒" '显示程序运行所花费的时间

End Sub

Sub A01_编多表()

    '程序编写：Frank
    '编写日期：2014年3月20日
    '程序功能：按指定要求对一组特定表格（11列足彩表）进行编辑
    
    Application.ScreenUpdating = False '关闭屏幕更新, 以提高代码运行速度
    st = VBA.Timer '计时器
    Dim TN As Integer '定义一个名tn的变量,用于计数表格数量
    On Error Resume Next '遇到错误忽略之
    
    TN = ActiveDocument.Tables.Count
    
    For i = 1 To TN
        ActiveDocument.Tables(i).Select
        A01_编单表
    Next i
      Application.ScreenUpdating = True '恢复屏幕更新
      MsgBox "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒" '显示程序运行所花费的时间
End Sub
Sub A01_newTextbox()
    Dim docNew As Document
    Dim newTextbox As Shape

    'Create a new document and add a text box
    Set docNew = Documents.Add
    Set newTextbox = docNew.Shapes.AddTextbox _
        (Orientation:=msoTextOrientationHorizontal, _
        Left:=100, Top:=100, Width:=50, Height:=25)

    'Add text to the text box
    newTextbox.TextFrame.TextRange = "扣分"
End Sub

Sub A01_批量处理表头()
    Dim N As Integer
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    N = ActiveDocument.Tables.Count
    If N > 1 Then
        For i = 1 To N
        ActiveDocument.Tables(i).Select
        Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.Expand Unit:=wdParagraph
        A1 = Selection.Range.Text
        If Len(A1) < 2 Then Selection.Delete
        Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.Expand Unit:=wdParagraph
        A1 = Selection.Paragraphs(1).Range.Text
        A1 = Trim(A1)
        If Left(A1, 2) = "单位" Then
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.Expand Unit:=wdParagraph
            If Len(Selection.Range.Text) < 2 Then Selection.Delete
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.Expand Unit:=wdParagraph
            Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
            B01_选定文字变为表头
        Else
            B01_选定文字变为表头
        End If
        Next i
    End If
End Sub

Sub A01_批量处理表头E()
    Dim N As Integer
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    N = ActiveDocument.Tables.Count
    If N > 1 Then
    For i = 1 To N
    ActiveDocument.Tables(i).Select
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.Expand Unit:=wdParagraph
    A1 = Selection.Range.Text
    If Len(A1) < 2 Then Selection.Delete
    Selection.MoveUp Unit:=wdLine, Count:=1
    Selection.Expand Unit:=wdParagraph
    A1 = Selection.Paragraphs(1).Range.Text
    A1 = Trim(A1)
    If Left(A1, 4) = "Unit" Then
        Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.Expand Unit:=wdParagraph
        If Len(Selection.Range.Text) < 2 Then Selection.Delete
        Selection.MoveUp Unit:=wdLine, Count:=1
        Selection.Expand Unit:=wdParagraph
        Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
        B01_选定文字变为表头E
    Else
        B01_选定文字变为表头E
    End If
    
    Next i
    
    End If
End Sub


Sub A01_批量加粗表格合计行()
    Dim N As Integer
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    N = ActiveDocument.Tables.Count
    If N > 1 Then
        For i = 1 To N
            ActiveDocument.Tables(i).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.SelectColumn
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .ClearFormatting
                .Execute FindText:="合"
                If .Found = True Then
                    .Parent.Expand Unit:=wdParagraph
                End If
            End With
            A1 = Selection.Range.Text
            P1 = InStr(1, A1, "计", 1)
            If P1 > 0 Then
                Selection.SelectRow
                Selection.Range.Font.Bold = True
            End If
        Next i
    End If
    
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    N = ActiveDocument.Tables.Count
    If N > 1 Then
        For i = 1 To N
            ActiveDocument.Tables(i).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.SelectColumn
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .ClearFormatting
                .Execute FindText:="总"
                If .Found = True Then
                    .Parent.Expand Unit:=wdParagraph
                End If
            End With
            A1 = Selection.Range.Text
            P1 = InStr(1, A1, "计", 1)
            If P1 > 0 Then
                Selection.SelectRow
                Selection.Range.Font.Bold = True
            End If
        Next i
    End If
    
End Sub

Sub A01_批量加粗表格合计行E()
    Dim N As Integer
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    N = ActiveDocument.Tables.Count
    If N > 1 Then
    For i = 1 To N
        ActiveDocument.Tables(i).Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.SelectColumn
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .ClearFormatting
            .Execute FindText:="Total"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
        End If
    End With
    A1 = Selection.Range.Text
    P1 = InStr(1, A1, "Total", 1)
    If P1 > 0 Then
        Selection.SelectRow
        Selection.Range.Font.Bold = True
    End If
    
    Next i
    
    End If
End Sub

Sub A01_批量加粗表格中的特定行()
    '表格中带有“一、”、“二、”、“三、”的行加粗，最多加粗至“二十”
    Dim A As Variant
    A = Array("一", "二", "三", "四", "五", "六", "七", "八", "九", "十" _
          , "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十")
    Selection.Tables(1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectColumn
    Set MyRange = Selection.Range
    
    For j = 0 To UBound(A)
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=A(j) & "、"
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
            Else
                Exit For
            End If
        End With
        Selection.SelectRow
        Selection.Range.Font.Bold = True
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        MyRange.Select
    Next j
    
    Selection.Tables(1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Tables(1).Cell(1, 2).Select
    Selection.SelectColumn
    Set MyRange = Selection.Range
    
    For j = 0 To UBound(A)
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=A(j) & "、"
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
            Else
                Exit For
            End If
        End With
        Selection.SelectRow
        Selection.Range.Font.Bold = True
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        MyRange.Select
    Next j
End Sub

Sub A01_批量加粗表格中的特定行E()
    '表格中带有“一、”、“二、”、“三、”的行加粗，最多加粗至“二十”
    Dim A As Variant
    A = Array("I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X" _
          , "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX")
    Selection.Tables(1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectColumn
    Set MyRange = Selection.Range
    
    For j = 0 To UBound(A)
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=A(j) & "."
            If .Found = True Then
                .Parent.Expand Unit:=wdParagraph
            Else
                Exit For
            End If
        End With
        Selection.SelectRow
        Selection.Range.Font.Bold = True
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        MyRange.Select
    Next j
End Sub


Sub A01_选定段落格式固定行距12磅()
    
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 12
    End With
    
End Sub

Sub A01_部门目录()
    
    Dim P1 As String
    Dim TN As Integer
    
    P1 = ActiveDocument.Paragraphs(1).Range.Text
    P1 = Left(P1, Len(P1) - 1)
    TN = ActiveDocument.Tables.Count
    
    If TN = 3 And P1 = "政府部门统计调查项目目录" Then
    
    网页页面
    
    ActiveDocument.Tables(1).Select
    A01_部门项目目录表格格式调整
    ActiveDocument.Tables(2).Select
    A01_部门项目目录表格格式调整
    ActiveDocument.Tables(3).Select
    A01_部门项目目录表格格式调整
    
    Else
    
    MsgBox "抱歉！本文档不符合执行本宏命令条件！"
    End If
    
End Sub

Sub A01_部门项目目录表格格式调整()

    表格B
    Selection.Tables(1).Select
    
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
    
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(1)
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectRow
    Selection.Font.Bold = wdToggle
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Shading.BackgroundPatternColor = wdColorGray10
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectColumn
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
End Sub

Sub A01_部门目录文档()
    Dim MyDOC As Document, MyDir As String
    Set MyDOC = Documents.Add
    网页页面
    For i = 1 To 10
        Selection.TypeParagraph
    Next i
    MyDOC.Range.Select
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 12
        .Alignment = wdAlignParagraphCenter
    End With
    With Selection.Font
        .NameFarEast = "宋体"
        .NameAscii = "宋体"
        .Size = 12
        .Bold = True
    End With
    MyDOC.Paragraphs(1).Range.Select
    Selection.TypeText Text:="政府部门统计调查项目目录"
    MyDOC.Paragraphs(1).Range.Select
    Selection.Font.Color = wdColorRed
    
    MyDOC.Paragraphs(3).Range.Select
    Selection.TypeText Text:="国家发展和改革委员会"
    MyDOC.Paragraphs(3).Range.Select
    Selection.Font.Color = wdColorBlue

    MyDOC.Paragraphs(5).Range.Select
    Selection.TypeText Text:="审批项目一览表"
    MyDOC.Paragraphs(5).Range.Select
    Selection.Font.Color = wdColorBlack
    
    MyDOC.Paragraphs(9).Range.Select
    Selection.TypeText Text:="备案项目一览表"
    MyDOC.Paragraphs(9).Range.Select
    Selection.Font.Color = wdColorBlack
    


    'Selection.Sections(1).Footers(1).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberOutside, FirstPage:=True
   ' MyDoc.SaveAs FileName:=FN, FileFormat:=wdFormatDocument    '保存文件

End Sub

Sub A01_星期替换()
'
' 全角数字转换S Macro
' 宏在 2003-6-27 由 DHG 录制
'
    Selection.WholeStory
    Dim A As Variant
    Dim B As Variant
    A = Array("一", "二", "三", "四", "五", "六", "日")
    B = Array("Mon", "Tue", "Wen", "Thu", "Fri", "Sat", "Sun")
    For i = 0 To 9
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = A(i)
        .Replacement.Text = B(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
End Sub

Sub A01_删空格回车()

    Dim doc As Document
    Dim RNG As Range
    Dim S1 As String, S2 As String, str As String
    Dim i As Integer

    Set doc = ActiveDocument
    Set RNG = Selection.Range
    'Set rng = doc.Range(start:=doc.Paragraphs(2).Range.start, End:=doc.Paragraphs(4).Range.End)
    S1 = RNG.Text
    str = ""
    For i = 1 To Len(S1)
        S2 = Mid(S1, i, 1)
        If S2 <> " " And S2 <> Chr(13) Then
            str = str & S2
        End If
    Next i
    MsgBox str
    
End Sub

Sub A01_不合并表头表注()

    N = ActiveDocument.Tables.Count
    If N > 0 Then
    For i = 1 To N
        ActiveDocument.Tables(i).Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.SelectRow
            If Left(Selection.Range.Text, 1) = "表" Then
                Selection.Cut
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.PasteSpecial DataType:=wdPasteText
            End If
        ActiveDocument.Tables(i).Select
        Selection.EndKey
        Selection.SelectRow
        If InStr(1, Selection.Range.Text, "注：", 1) > 0 Then
            Selection.Cut
            Selection.PasteSpecial DataType:=wdPasteText
        End If
    Next i
    End If

    N = ActiveDocument.Tables.Count
    If N > 0 Then
    For i = 1 To N
        ActiveDocument.Tables(i).Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.SelectRow
            If InStr(1, Selection.Range.Text, "单位：", 1) > 0 Then
                Selection.Cut
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.TypeParagraph
                Selection.PasteSpecial DataType:=wdPasteText
            End If
    Next i
    End If

End Sub

Sub A01_不合并表头表注E()

    N = ActiveDocument.Tables.Count
    If N > 0 Then
    For i = 1 To N
        ActiveDocument.Tables(i).Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.SelectRow
            If Left(Selection.Range.Text, 5) = "Table" Then
                Selection.Cut
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.PasteSpecial DataType:=wdPasteText
            End If
        ActiveDocument.Tables(i).Select
        Selection.EndKey
        Selection.SelectRow
        If InStr(1, Selection.Range.Text, "Note:", 1) > 0 Then
            Selection.Cut
            Selection.PasteSpecial DataType:=wdPasteText
        End If
    Next i
    End If

    N = ActiveDocument.Tables.Count
    If N > 0 Then
    For i = 1 To N
        ActiveDocument.Tables(i).Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.SelectRow
            If InStr(1, Selection.Range.Text, "Unit:", 1) > 0 Then
                Selection.Cut
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.TypeParagraph
                Selection.PasteSpecial DataType:=wdPasteText
            End If
    Next i
    End If
    
     Selection.HomeKey Unit:=wdStory

End Sub

Sub A01_设置快捷键()

    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyV, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="粘贴文本"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyK, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="中文空格"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyZ, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="A00_网页格式"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyQ, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="标准网页格式W"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyB, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="表格B"
        
End Sub

Sub A01_尾注编号()

    Dim RNG As Range
    Set RNG = ActiveDocument.Range
    TXT = RNG.Text
    TXT = Left(TXT, Len(TXT) - 1)
    L = Len(TXT)
    C = "[ ]"
    i = 1
    P = InStr(1, TXT, C, 1)
    
    Application.ScreenUpdating = False '关闭屏幕更新
    
    Do While P > 0
    
    If P > 0 Then
        S1 = Left(TXT, P - 1)
        S2 = Right(TXT, L - P - 2)
        RNG.Text = S1 & "[" & i & "]" & S2
    End If
    
    RNG = ActiveDocument.Range
    TXT = RNG.Text
    TXT = Left(TXT, Len(TXT) - 1)
    L = Len(TXT)
    C = "[ ]"
    i = i + 1
    P = InStr(1, TXT, C, 1)
    
    If P = 0 Then Exit Do
    
    Loop
    
    Selection.HomeKey Unit:=wdStory '插入点置于文档开始处
    Application.ScreenUpdating = True '恢复屏幕更新

End Sub
Sub A01_删除全部尾注()

    Dim N As Integer
    N = ActiveDocument.Endnotes.Count
    
    If N > 0 Then
        For Each nt In ActiveDocument.Endnotes
            nt.Delete
        Next nt
    End If

End Sub


Sub A01_显示尾注内容()

    Dim N As Integer
    N = ActiveDocument.Endnotes.Count
    MsgBox N
    
    If N > 0 Then
        For Each nt In ActiveDocument.Endnotes
            MsgBox nt.Range.Text
        Next nt
    End If

End Sub

Sub A01_图片版式由浮动型转换为嵌入型()

    On Error Resume Next
    Dim N As Integer
    Application.ScreenUpdating = False '关闭屏幕更新
    N = ActiveDocument.Shapes.Count
    Selection.HomeKey Unit:=wdStory
    If N > 0 Then
        For Each ishape In ActiveDocument.Shapes
            ishape.Select
            ishape.ConvertToInlineShape
        Next ishape
    End If
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
    MsgBox "共转换了" & N & "张图片为嵌入型"

End Sub

Sub A01_图片版式由嵌入型转换为浮动型()
    On Error Resume Next
    Dim N As Integer
    Application.ScreenUpdating = False '关闭屏幕更新
    Selection.HomeKey Unit:=wdStory
    N = ActiveDocument.InlineShapes.Count
    If N > 0 Then
        For Each ishape In ActiveDocument.InlineShapes
            ishape.Select
            ishape.ConvertToShape
        Next ishape
    End If
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '恢复屏幕更新
    MsgBox "共转换了" & N & "张图片为浮动四周型"
    
End Sub

Sub A01_测试宏命令2()  '功能：调试宏命令

 A01_删除千分位逗号
 
'A01_图片版式由浮动型转换为嵌入型
'A01_图片版式转换
    
End Sub


Sub A01_取消行首缩进()
    
    With Selection.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0.1)
        .RightIndent = CentimetersToPoints(0.1)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphJustify
        .WidowControl = True
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
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = False
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
End Sub
Sub A01_删除千分位逗号()

    Dim MyRange As Range
    Dim T As String, S1 As String, S2 As String, S3 As String, S4 As String, TT As String
    Dim P As Integer, N As Integer, i As Integer, j As Integer
    
    Application.ScreenUpdating = False '关闭屏幕更新
    
    '设定程序作用范围为用户选定的区域，如果用户没有选定区域，则默认选定区域为整篇文档
    If Len(Selection.Range.Text) = 0 Then Selection.WholeStory
    Set MyRange = Selection.Range
    t1 = "，" '指定字符，这里是逗号
    TT = ""  '定义一个空字符串
    st = VBA.Timer '计时器
    j = 0 '初始化计数器
    N = MyRange.Paragraphs.Count
    
        For i = 1 To N
            T = MyRange.Paragraphs(i).Range.Text
            T = Left(T, Len(T) - 1)  '段落文本，不含换行符（回车）
            P = InStr(1, T, t1, 1)   '指定字号在段落中的位置
            TT = ""
            
            If P > 0 Then
            Do Until P = 0
                If InStr(1, T, t1, 1) > 0 Then
                    P = InStr(1, T, t1, 1)
                    S1 = Left(T, P - 1)  ' 指定字符前面的文本
                    S2 = Right(T, Len(T) - P)   ' 指定字符后面的文本
                    S3 = Right(S1, 1) ' 指定字符前面一个字符
                    S4 = Left(S2, 1)  ' 指定字符后面一个字符
                    '如果指定字符前面一个字符和后面一个字符均为数字，则删除该逗号
                    If Asc(S3) > 45 And Asc(S3) < 58 And Asc(S4) > 45 And Asc(S4) < 58 Then
                        TT = TT & S1
                        j = j + 1
                    Else
                        TT = TT & S1 & t1
                    End If
                    T = S2
                    P = InStr(1, T, t1, 1)
                Else
                    Exit Do
                End If
            Loop
            MyRange.Paragraphs(i).Range.Text = TT & S2 & Chr(13)
            End If
            
        Next i
    Selection.HomeKey Unit:=wdStory '插入点置于文档开始处
    Application.ScreenUpdating = True '恢复屏幕更新
    MsgBox "共替换  " & j & "  处。" & Chr(13) & "程序运行用时" & VBA.Format(Timer - st, "0.00") & "秒" '显示程序运行所花费的时间

End Sub

Sub A01_插入六角方括号()
    Selection.TypeText Text:=Chr(-24142) & Year(Date) & Chr(-24141)
End Sub

Sub A01_拆分文档()

    Application.ScreenUpdating = False '关闭屏幕更新

    Dim NM As String   'NM -- 文件名
    Dim N As Integer   'N -- 文件数
    Dim TXT As String  'TXT -- 文本
    Dim FN() As Variant  '定义数组
    Dim TB As Table  '定义表格变量
    Dim RS As Range   '开始点
    Dim RE As Range   '结束点
    Dim MyRange As Range '需要拷贝的区域
    
    ChangeFileOpenDirectory "D:\01 MyFiles"  '设置打开文件的路径
    Documents.Open ("目录.doc")
    Set TB = ActiveDocument.Tables(1)
    TXT = TB.Cell(1, 1).Range.Text
    TXT = Left(TXT, Len(TXT) - 1)
    ReDim Preserve FN(1)  '重新定义数组FN
    FN(0) = TXT  '初始化数组元素
    For i = 2 To TB.Rows.Count
        TXT = TB.Cell(i, 1).Range.Text
        TXT = Left(TXT, Len(TXT) - 1)
        ReDim Preserve FN(UBound(FN) + 1)
        FN(UBound(FN) - 1) = TXT
    Next i
    
    'MsgBox FN(UBound(FN) - 1)
   ' MsgBox UBound(FN)
    
    Documents.Close SaveChanges:=wdDoNotSaveChanges  '关闭目录文档
    
    Documents.Open ("测试文档.doc")  '打开测试文档
    For i = 0 To UBound(FN) - 1
        NM = FN(i)
        NM1 = FN(i + 1)
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = NM
    End With
    Selection.Find.Execute
    
    Set RS = Selection.Range
    
    Selection.MoveDown Unit:=wdLine, Count:=1
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = NM1
    End With
    Selection.Find.Execute
    Selection.MoveUp Unit:=wdLine, Count:=1
    Set RE = Selection.Range
    
    If i = UBound(FN) - 1 Then
        Set MyRange = ActiveDocument.Range(start:=RS.start, End:=ActiveDocument.Range.End - 1)
    Else
        Set MyRange = ActiveDocument.Range(start:=RS.start, End:=RE.start)
    End If
    
    MyRange.Copy
    Documents.Add
    Selection.Paste
    
    If Len(i + 1) = 1 Then
        NM = "A0" & i + 1
    Else
        NM = "A" & i + 1
    End If
    
    ActiveDocument.SaveAs FileName:=NM & ".doc", FileFormat:=wdFormatDocument
    ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
    
    Next i
    
    ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
    Application.ScreenUpdating = True '打开屏幕更新
    ChangeFileOpenDirectory "D:\00 F2013\" '恢复打开文件的缺省路径
    
End Sub


Sub A00_浮动图片转换为嵌入式图片()
    
    N = ActiveDocument.InlineShapes.Count ' 取得文档中图片数
    MsgBox N
    
    N_SHP = ActiveDocument.Shapes.Count
    
    MsgBox N_SHP
    
    '如果文档中有浮动式图片，则将其转换为嵌入式图片
    If ActiveDocument.Shapes.Count > 0 Then
        For Each oShape In ActiveDocument.Shapes
            oShape.ConvertToInlineShape
        Next
    End If
    
    N = ActiveDocument.InlineShapes.Count ' 取得文档中图片数
    MsgBox N
    
End Sub
Sub 纯文本粘贴()
'
' 纯文本粘贴 宏
'
'
    CommandBars("Office Clipboard").Visible = False
    Selection.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:= _
        wdInLine, DisplayAsIcon:=False
    Selection.TypeParagraph
    Selection.TypeParagraph
End Sub


Sub G00_每日调查()

    Dim MyDOC As Document
    Dim MyRange As Range
    
    Application.ScreenUpdating = False '关闭屏幕更新
    On Error Resume Next
    
    '创建一个新文档
    Set MyDOC = Documents.Add
    
    '将新文档设置为公文页面
    G00_公文页面设置
    
    '插入15个空行
    For i = 1 To 13
        Selection.TypeParagraph
    Next i
    
    '设置段落格式、字体字号
    MyDOC.Range.Select
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    Selection.ParagraphFormat.LineSpacing = 30
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.name = "仿宋_GB2312"
    Selection.Font.Size = 15
    
    '设置第一行段落格式、字体字号
    Set MyRange = MyDOC.Paragraphs(1).Range
    MyRange.Text = "未经许可"
    MyRange.Font.name = "黑体"
    MyRange.Font.Size = 16
    MyRange.ParagraphFormat.Alignment = wdAlignParagraphRight
    
     '设置第二行段落格式、字体字号
    Set MyRange = MyDOC.Paragraphs(2).Range
    MyRange.Text = "不得转载"
    MyRange.Font.name = "黑体"
    MyRange.Font.Size = 16
    MyRange.ParagraphFormat.Alignment = wdAlignParagraphRight
    
    '设置文头大字
    Set MyRange = MyDOC.Paragraphs(3).Range
    MyRange.Text = "每日调查"
    MyRange.Font.name = "华文行楷"
    MyRange.Font.Size = 80
    MyRange.Font.Color = wdColorRed
    MyRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    MyRange.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    MyRange.ParagraphFormat.SpaceBefore = 20
    MyRange.ParagraphFormat.LineUnitBefore = 4
    
    '设置文号
    Set MyRange = MyDOC.Paragraphs(4).Range
    MyRange.Text = "（" & Year(Date) & "年第 期）"
    MyRange.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    MyRange.ParagraphFormat.SpaceBefore = 15
    MyRange.ParagraphFormat.SpaceAfter = 15
    MyRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    '设置发文单位和发文日期
    Set MyRange = MyDOC.Paragraphs(5).Range
    MyRange.ParagraphFormat.TabStops.ClearAll
    MyRange.Select
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(15.2), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    Selection.TypeText Text:="国家统计局办公室" & vbTab & Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日"
    MyRange.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    MyRange.ParagraphFormat.Alignment = wdAlignParagraphLeft
    MyRange.Select
    'Set MyRange = ActiveDocument.Range(start:=Selection.Range.start, End:=ActiveDocument.Range.End - 1)
    Set MyRange = ActiveDocument.Range(start:=Selection.Characters(1).start, End:=Selection.Characters(8).End)
    MyRange.Select
    Selection.Font.name = "黑体"
    
    '设置红线
    Set MyRange = MyDOC.Paragraphs(6).Range
    MyRange.Borders(wdBorderTop).LineStyle = Options.DefaultBorderLineStyle
    MyRange.Borders(wdBorderTop).LineWidth = wdLineWidth300pt
    MyRange.Borders(wdBorderTop).Color = wdColorRed
    
     '设置标题段落格式、字体字号
    Set MyRange = MyDOC.Paragraphs(8).Range
    MyRange.Font.name = "方正小标宋_GBK"
    MyRange.Font.Size = 22
    MyRange.Text = "标题"
    MyRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
     '设置正文段落格式、字体字号
    Set MyRange = MyDOC.Paragraphs(9).Range
    MyRange.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0.35)
    MyRange.ParagraphFormat.CharacterUnitFirstLineIndent = 2
    MyRange.Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeParagraph
    Application.ScreenUpdating = True '恢复屏幕更新

End Sub
Sub G00_公文页面设置()
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

Sub 我的宏1()
'
' 我的宏1 宏
'
'
    Selection.TypeText Text:="国家统计局"
    Selection.TypeParagraph
End Sub

Sub GetDocumentName()
    Dim strDocName As String
    strDocName = ActiveDocument.name
    MsgBox strDocName
End Sub


Sub PrintThreePages()
    ActiveDocument.PrintOut Range:=wdPrintRangeOfPages, Pages:="1-3"
End Sub

Sub CloseDocument()
    Documents(1).Close
End Sub

Sub SetBoldRange()
    Dim rngDoc As Range
    Set rngDoc = ActiveDocument.Range(start:=0, End:=10)
    rngDoc.Bold = True
End Sub

Sub BoldRange()
    ActiveDocument.Range(start:=0, End:=10).Bold = True
End Sub

Sub 指定范围并加粗居中()
    Dim RNG As Range
    Set RNG = ActiveDocument.Paragraphs(1).Range
    With RNG
        .Bold = True
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Font
            .name = "黑体"
            .Size = 15
        End With
    End With
End Sub

Sub SelectFirstWord()
    ActiveDocument.Words(1).Select
    Selection.Font.Bold = True
End Sub

Sub ReturnCellContentsToArray()
    Dim intCells As Integer
    Dim celTable As Cell
    Dim strCells() As String
    Dim intCount As Integer
    Dim rngText As Range
    
    If ActiveDocument.Tables.Count >= 1 Then
        With ActiveDocument.Tables(1).Range
            intCells = .Cells.Count
            ReDim strCells(intCells)
            intCount = 1
            For Each celTable In .Cells
                Set rngText = celTable.Range
                rngText.MoveEnd Unit:=wdCharacter, Count:=-1
                strCells(intCount) = rngText
                intCount = intCount + 1
            Next celTable
        End With
    End If
End Sub

Sub 判断文档是否打开()

    Dim doc As Document
    Dim docFound As Boolean
    
    For Each doc In Documents
        If InStr(1, doc.name, "sample.doc", 1) Then
            doc.Activate
            docFound = True
            Exit For
        Else
            docFound = False
        End If
    Next doc
    
    If docFound = False Then Documents.Open FileName:="Sample.doc"
    
End Sub


Sub A00_解除文档保护属性()

    Dim doc As Document
    Set doc = ActiveDocument
    If doc.ProtectionType <> wdNoProtection Then doc.Unprotect
    
End Sub

Sub A00_加文档保护属性()

    Dim doc As Document
    Set doc = ActiveDocument
    If doc.ProtectionType = wdNoProtection Then doc.Protect Type:=wdAllowOnlyFormFields
    
End Sub
Sub A00_接受文档所有修订()

    Dim doc As Document
    Set doc = ActiveDocument
    doc.AcceptAllRevisions
    
End Sub

Sub 大写()
'
' 大写 宏
'
'

End Sub
