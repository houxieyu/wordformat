Attribute VB_Name = "��ҳ"
Sub ��׼��ҳ��ʽW()

    'st = VBA.Timer '�������м�ʱ��
    Application.ScreenUpdating = False '�ر���Ļ����
    
    On Error Resume Next
    
    Selection.HomeKey Unit:=wdStory
     A01_������滻
    
    '��ԭ�ĵ�ȫ������ת��Ϊ���ı���ʽ
    Selection.WholeStory
    Selection.Cut
    Selection.Collapse Direction:=wdCollapseStart
    
    CommandBars("Office Clipboard").Visible = False
    Selection.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:=wdInLine, DisplayAsIcon:=False
    
    'Selection.Range.PasteSpecial DataType:=wdPasteText
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    
    Application.Run MacroName:="ɾ��ǰ��"      'Ӧ�ú����ɾ��ǰ��
    Application.Run MacroName:="ɾ����"        'Ӧ�ú����ɾ����
    Application.Run MacroName:="ȫ��ת��"      'Ӧ�ú����ȫ��ת��
    Application.Run MacroName:="��ǰ�ӿ�"      'Ӧ�ú������ǰ�ӿ�
    
    '���ĵ���ʽ��ΪС���Σ�ȥ����������
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
    
    '���ĵ��е�һ�������Ϊ����

    Selection.WholeStory
    Dim A As Variant
        A = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ" _
             , "ʮһ", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "��ʮ")
    For j = 0 To 19
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
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Paragraphs(1).Range.Select
    Selection.Delete Unit:=wdCharacter, Count:=1
    
    'ÿ��������һ����
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
        A01_������滻1
        Selection.HomeKey Unit:=wdStory
        
    Application.ScreenUpdating = True '�ָ���Ļ����
    'MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��" '��ʾ�������е�ʱ��

    'CommandBars("Office Clipboard").Visible = True

End Sub
Sub ��׼��ҳ��ʽW1()
    
    'st = VBA.Timer '�������м�ʱ��
    Application.ScreenUpdating = False '�ر���Ļ����
    
    On Error Resume Next

    Selection.HomeKey Unit:=wdStory
    A01_������滻

    Selection.WholeStory
    Selection.Cut
    Selection.Collapse Direction:=wdCollapseStart
    
    CommandBars("Office Clipboard").Visible = False
    Selection.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:=wdInLine, DisplayAsIcon:=False
    
    'Selection.Range.PasteSpecial DataType:=wdPasteText
    
   'ɾ������ǰ�Ͷ����Ŀո�
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

    Dim O As Variant
    Dim R As Variant
    O = Array("^l", "  ", "^p^p", "^p", "����^p")
    R = Array("^p", " ", "^p", "^p����", "")
    
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
A = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ" _
          , "ʮһ", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "��ʮ")
    
    For j = 0 To 19
    
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
    
    Application.Run MacroName:="ȫ��ת��"      'Ӧ�ú����ȫ��ת��

    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Paragraphs(1).Range.Select
    Selection.Delete Unit:=wdCharacter, Count:=1
    Selection.WholeStory
    Selection.EndKey Unit:=wdStory
    Selection.TypeBackspace
    Selection.HomeKey Unit:=wdStory
    A01_������滻1
    Selection.HomeKey Unit:=wdStory
        
    Application.ScreenUpdating = True '�ر���Ļ����
   ' MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��" '��ʾ�������е�ʱ��
    
   CommandBars("Office Clipboard").Visible = True
     
End Sub

