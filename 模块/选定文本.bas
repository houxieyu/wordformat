Attribute VB_Name = "选定文本"
Sub B01_选定文字变为表头()
    
    TT = Selection.Range.Text
    TT = Trim(TT)
    Set MyRange = Selection.Range
    N = Selection.Paragraphs.Count
    If N = 2 Then
        B01_选定文字变为表头2
    Else
        If Len(TT) > 30 Then
            Exit Sub
        Else
            TT = Left(TT, Len(TT) - 1)
            P = InStr(1, TT, " ", 1)
            If P > 0 Then
                TT = Left(TT, P - 1) & "　" & Right(TT, Len(TT) - P)
            End If
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
                Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
                Selection.Shading.Texture = wdTextureNone
                Selection.Shading.ForegroundPatternColor = wdColorAutomatic
                Selection.Shading.BackgroundPatternColor = wdColorAutomatic
                Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                Selection.Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
                Selection.SelectRow
                Selection.Font.Size = 12
                Selection.Font.Bold = True
                Selection.ParagraphFormat.LineSpacing = 16
                Selection.Rows.HeightRule = wdRowHeightAtLeast
                Selection.Rows.Height = CentimetersToPoints(1)
            End If
        End If
    End If
    MyRange.Select
    Selection.Delete
    Selection.TypeBackspace
    
End Sub

Sub B01_选定文字变为表头E()
    
    TT = Selection.Range.Text
    TT = Trim(TT)
    Set MyRange = Selection.Range
    N = Selection.Paragraphs.Count
    If N = 2 Then
        B01_选定文字变为表头2E
    Else
        If Len(TT) < 1 Then
            MsgBox "请选择文本！"
        Else

        TT = Left(TT, Len(TT) - 1)
        P = InStr(1, TT, " ", 1)
        If P > 0 Then
            TT = Left(TT, P - 1) & " " & Right(TT, Len(TT) - P)
        End If
    
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
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = wdColorAutomatic
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
    Selection.Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
    Selection.SelectRow
    Selection.Font.Size = 10.5
    Selection.Font.Bold = True
    Selection.ParagraphFormat.LineSpacing = 16
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(1)
    
    Else
        MsgBox "请重新选定文本!"
    End If
    End If
    End If
    MyRange.Select
    Selection.Delete
    
End Sub

Sub B01_选定文字变为表头2()
    
    TT = Selection.Range.Text
    N = Selection.Paragraphs.Count
    TT = Left(TT, Len(TT) - 1)
    
    If N = 2 Then
        TT1 = Selection.Paragraphs(1).Range.Text
        TT1 = Trim(TT1)
        TT1 = Left(TT1, Len(TT1) - 1)
        P = InStr(1, TT1, " ", 1)
        If P > 0 Then
            TT1 = Left(TT1, P - 1) & "　" & Right(TT1, Len(TT1) - P)
        End If
        TT2 = Selection.Paragraphs(2).Range.Text
        TT2 = Trim(TT2)
        TT2 = Left(TT2, Len(TT2) - 1)
    End If
    
    Selection.MoveDown Unit:=wdLine, Count:=1
    
    If Selection.Information(wdWithInTable) = True Then
        Set myrng = Selection.Tables(1).Range
        Set MyTab = Selection.Tables(1)
    MyTab.Cell(1, 1).Select
    Selection.SelectRow
    Selection.InsertRowsAbove 2
    
    MyTab.Cell(1, 1).Select
    Selection.SelectRow
    Selection.Cells.Merge
    Selection.Cells(1).Range.Text = TT1
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = wdColorAutomatic
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.SelectRow
    Selection.Font.Bold = True
    Selection.Font.Size = 12
    Selection.ParagraphFormat.LineSpacing = 16
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(1)
    
    MyTab.Cell(2, 1).Select
    Selection.SelectRow
    Selection.Cells.Merge
    Selection.Cells(1).Range.Text = TT2
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
    Selection.Range.Font.Bold = False
    Selection.Rows.Height = CentimetersToPoints(0.5)
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    MyTab.Cell(1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    Else
        MsgBox "请重新选定文本!"
    End If
        
End Sub

Sub B01_选定文字变为表头2E()
    
    TT = Selection.Range.Text
    N = Selection.Paragraphs.Count
    TT = Left(TT, Len(TT) - 1)
    
    If N = 2 Then
        TT1 = Selection.Paragraphs(1).Range.Text
        TT1 = Trim(TT1)
        TT1 = Left(TT1, Len(TT1) - 1)
        P = InStr(1, TT1, " ", 1)
        If P > 0 Then
            TT1 = Left(TT1, P - 1) & " " & Right(TT1, Len(TT1) - P)
        End If
        TT2 = Selection.Paragraphs(2).Range.Text
        TT2 = Trim(TT2)
        TT2 = Left(TT2, Len(TT2) - 1)
    End If
    
    Selection.MoveDown Unit:=wdLine, Count:=1
    
    If Selection.Information(wdWithInTable) = True Then
        Set myrng = Selection.Tables(1).Range
        Set MyTab = Selection.Tables(1)
    MyTab.Cell(1, 1).Select
    Selection.SelectRow
    Selection.InsertRowsAbove 2
    
    MyTab.Cell(1, 1).Select
    Selection.SelectRow
    Selection.Cells.Merge
    Selection.Cells(1).Range.Text = TT1
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = wdColorAutomatic
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.SelectRow
    Selection.Font.Bold = True
    Selection.Font.Size = 10.5
    Selection.ParagraphFormat.LineSpacing = 16
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(1)
    
    MyTab.Cell(2, 1).Select
    Selection.SelectRow
    Selection.Cells.Merge
    Selection.Cells(1).Range.Text = TT2
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
    Selection.Range.Font.Bold = False
    Selection.Rows.Height = CentimetersToPoints(0.5)
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    MyTab.Cell(1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    Else
        MsgBox "请重新选定文本!"
    End If
        
End Sub

Sub B01_选定文本变为表注()
    
    TT = Selection.Range.Text
    Selection.Delete

    If Len(TT) < 1 Then
        MsgBox "请选择文本！"
    Else
        TT = Left(TT, Len(TT) - 1)
        TT = Trim(TT)
        Selection.MoveUp Unit:=wdLine, Count:=1
    If Selection.Information(wdWithInTable) = True Then
        Selection.SelectRow
        Selection.InsertRowsBelow 1
        Selection.Cells.Merge
        Selection.Cells(1).Range.Text = TT
        Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        Selection.Borders(wdBorderTop).LineWidth = wdLineWidth150pt
        Selection.Shading.Texture = wdTextureNone
        Selection.Shading.ForegroundPatternColor = wdColorAutomatic
        Selection.Shading.BackgroundPatternColor = wdColorAutomatic
        Selection.SelectRow
        Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        Selection.Font.NameFarEast = "楷体"
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Else
        MsgBox "请重新选定文本！"
    End If
    
    End If
    
End Sub

Sub B01_选定文本变为表注E()
    
    TT = Selection.Range.Text
    Selection.Delete

    If Len(TT) < 1 Then
        MsgBox "请选择文本！"
    Else
        TT = Left(TT, Len(TT) - 1)
        Selection.MoveUp Unit:=wdLine, Count:=1
    If Selection.Information(wdWithInTable) = True Then
        Selection.SelectRow
        Selection.InsertRowsBelow 1
        Selection.Cells.Merge
        Selection.Cells(1).Range.Text = TT
        Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
        Selection.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        Selection.Borders(wdBorderTop).LineWidth = wdLineWidth150pt
        Selection.Shading.Texture = wdTextureNone
        Selection.Shading.ForegroundPatternColor = wdColorAutomatic
        Selection.Shading.BackgroundPatternColor = wdColorAutomatic
        Selection.SelectRow
        Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    With Selection.Font
        .NameFarEast = "楷体"
        .NameAscii = "Arial"
        .NameOther = "Arial"
        .Size = 9
        .Italic = True
    End With
    
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Else
        MsgBox "请重新选定文本！"
    End If
    
    End If
    
End Sub
