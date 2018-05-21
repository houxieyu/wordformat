Attribute VB_Name = "字符处理"
Sub 删除空格K()
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " "
        .Replacement.Text = ""
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub A01_间隔号替换()
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "·"
        .Replacement.Text = "JGH"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub A01_间隔号替换1()
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "JGH"
        .Replacement.Text = "·"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub A01_间隔号替换2()
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(41380)
        .Replacement.Text = "JGH"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 空格替换()
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(32) & Chr(32)
        .Replacement.Text = Chr(-24159)
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub 空格替换1()
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(32)
        .Replacement.Text = Chr(-24159)
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub


Sub 删除中文空格()
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "　"
        .Replacement.Text = ""
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub 删段前空()
    Dim MyRange As Range
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If
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
End Sub
Sub 删段后空()
    Dim MyRange As Range
    If Len(Selection.Range.Text) = 0 Then Selection.WholeStory
    Set MyRange = Selection.Range
    If MyRange.Tables.Count > 0 Then Exit Sub
    Selection.ClearFormatting
    With MyRange
        pn = MyRange.Paragraphs.Count
        For i = 1 To pn
            T = MyRange.Paragraphs(i).Range.Text
            MyRange.Paragraphs(i).Range.Text = RTrim(Left(T, Len(T) - 1)) & Right(T, 1)
        Next i
     End With
    Selection.WholeStory
    Selection.EndKey Unit:=wdStory
    Selection.Delete Unit:=wdCharacter, Count:=1
End Sub

Sub 段前加空()
    Dim MyRange As Range  '定义一个范围变量
    '如果没有选择范围，则指定范围为整个文档
    If Len(Selection.Range.Text) = 0 Then Selection.WholeStory
    Set MyRange = Selection.Range '设定范围变量为选择的范围
    If MyRange.Tables.Count > 0 Then Exit Sub
    With MyRange
        pn = MyRange.Paragraphs.Count
        For i = 1 To pn
            T = Trim(MyRange.Paragraphs(i).Range.Text)
            S1 = "　　" & T
            MyRange.Paragraphs(i).Range.Text = S1
        Next i
     End With
        
End Sub

Sub 加空行()
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
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
    Selection.TypeBackspace
End Sub

Sub 删空行()

    Dim kh
    kh = True
    Do While kh = True
        Selection.WholeStory
        Selection.HomeKey Unit:=wdStory
        Selection.WholeStory
        With Selection.Find
        .ClearFormatting
        .Execute FindText:="^p^p"
        If .Found = True Then
        kh = True
        Else
        kh = False
        Exit Do
        End If
    End With

        
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        
        With Selection.Find
        .ClearFormatting
        .Execute FindText:="^p^p"
        If .Found = True Then
        kh = True
        Else
        kh = False
        Exit Do
        End If
    End With
    
    Loop
        
        Selection.WholeStory
        Selection.EndKey Unit:=wdStory

    Dim pn, T
    pn = ActiveDocument.Paragraphs.Count
    Set aRange = ActiveDocument.Range(start:=ActiveDocument.Paragraphs(pn).Range.start, End:=ActiveDocument.Paragraphs(pn).Range.End)
    T = aRange.Text
        If Len(T) < 2 Then
            aRange.Select
            Selection.Delete
        End If
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    Selection.MoveLeft Unit:=wdCharacter, Count:=1


End Sub
Sub 缩空()
    Dim MyRange As Range
        If Len(Selection.Range.Text) = 0 Then Selection.WholeStory
    Set MyRange = Selection.Range
    With MyRange
        pn = MyRange.Paragraphs.Count
        For i = 1 To pn
            T = Trim(MyRange.Paragraphs(i).Range.Text)
            L = Len(T)
            S2 = T
            p2 = InStr(1, S2, "  ", 1)
            Do Until p2 = 0
                If InStr(1, S2, "  ", 1) > 0 Then
                    S2 = Left(S2, p2 - 1) & Right(S2, Len(S2) - p2)
                End If
                p2 = InStr(1, S2, "  ", 1)
            Loop
            MyRange.Paragraphs(i).Range.Text = Trim(S2)
        Next i
    End With

End Sub
