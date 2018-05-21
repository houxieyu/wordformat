Attribute VB_Name = "A00"

Sub A00_ͼƬ���������ĵ�()

    st = VBA.Timer '�������м�ʱ��

    ChangeFileOpenDirectory "F:\DOC-��ͼ��"

    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next '���Դ���

    '����ĵ����и���ʽͼƬ������ת��ΪǶ��ʽͼƬ
    If ActiveDocument.Shapes.Count > 0 Then
        For Each oShape In ActiveDocument.Shapes
            oShape.ConvertToInlineShape
        Next
    End If

    'A01_���ϲ���ͷ��ע

    ' ��������������NM -- �ļ����� CN -- ͼƬ��  TN--�����
    Dim NM As String
    Dim CN As Integer
    Dim TN As Integer
        
    ' ȡ�õ�ǰ�ļ���
    Set MyDOC = Application.ActiveWindow.Document 'ָ��Ҫ������ĵ�ΪMyDoc
    NM = Left(MyDOC, Len(MyDOC) - 4)
    CN = ActiveDocument.InlineShapes.Count ' ȡ���ĵ���ͼƬ��
    TN = ActiveDocument.Tables.Count ' ȡ���ĵ��б����

    ' ��ԭ�ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
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
    ' �½�һ���հ��ĵ�, ������ʱ�洢ͼƬ��ָ��ΪDOC_CN
    If CN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_CN = Application.ActiveWindow.Document
        
        ҳ������ 'ִ�С�ҳ�����á�������
        
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
        
        ' �����ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
        DOC_CN.Activate
        For j = 1 To CN
            DOC_CN.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
        
        DOC_CN.Activate
        ActiveDocument.SaveAs FileName:=NM & "_ͼƬ.doc", FileFormat:=wdFormatDocument
        
    End If
    
End Sub

Sub A00_��׼��ҳ��ʽOLD()

    ' ��׼��ҳ��ʽ Macro
    
    Dim O As Variant
    Dim R As Variant
    Dim A As Variant
    Dim C As Variant
    Dim D As Variant
    Dim CC As Variant
    Dim DD As Variant
    
    On Error Resume Next
    Application.ScreenUpdating = False '�ر���Ļ����
    
    'O = Array("^l", "  ", "^p^p", "^p", "����^p", " ^p")
    'R = Array("^p", "��", "^p", "^p����", "", "^p")
    
    O = Array("^l", Chr(32) & Chr(32), "^p^p", "^p", "����^p", Chr(32) & Chr(13))
    R = Array("^p", Chr(-24159), "^p", "^p����", "", Chr(13))
    A = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ", _
          , "ʮһ", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "��ʮ", _
          "��ʮһ", "��ʮ��", "��ʮ��", "��ʮ��", "��ʮ��", "��ʮ��", "��ʮ��", "��ʮ��", "��ʮ��", "��ʮ")
    C = Array("��", "��", "��", "��", "��", "��", "��", "��", "��", "��", ",", ";", "��", "?", ":", "(", ")")
    D = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "��", "��", "%", ".", "��", "��", "��")
    CC = Array("^l", ",", ";", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", _
        "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", _
        "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", _
        "��", "��", "��", "��", "��", "��", "��")
    DD = Array("^p", "��", "��", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
        "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", _
        "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")
    
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
    
    '-------------------------------------------------
    '��ǡ�:���滻Ϊȫ�ǡ�����
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = ":^p"
            .Replacement.Text = "��^p"
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    '----------------------------------------------------
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    
    '��һ�����Ϊ���⣬��Ӵ־���
    Selection.Paragraphs(1).Range.Select
        s = Selection.Paragraphs(1).Range.Text
        If Len(s) < 30 Then
            Selection.Font.Bold = True
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
        
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



Sub A00_��ͼ��()

    '�趨�ĵ�����Ŀ¼
    Dim FD1 As String
    Dim FD2 As String
    
    FD1 = "C:\Users\zlzx-dhg\Desktop\00 OK_DOC"
    FD2 = "D:\00 OK_DOC"
    st = VBA.Timer '�������м�ʱ��

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

    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next '���Դ���

    '����ĵ����и���ʽͼƬ������ת��ΪǶ��ʽͼƬ
    If ActiveDocument.Shapes.Count > 0 Then
        For Each oShape In ActiveDocument.Shapes
            oShape.ConvertToInlineShape
        Next
    End If

    If ActiveDocument.Tables.Count > 0 Then
        A00_ȥ��ͷ��ע
    End If
    Selection.HomeKey Unit:=wdStory
    A00_ɾ����ҳ�ո�
    Selection.HomeKey Unit:=wdStory

    ' ��������������NM -- �ļ����� CN -- ͼƬ��  TN--�����
    Dim NM As String
    Dim CN As Integer
    Dim TN As Integer
        
    ' ȡ�õ�ǰ�ļ���
    Set MyDOC = Application.ActiveWindow.Document 'ָ��Ҫ������ĵ�ΪMyDoc
    If InStr(1, MyDOC, "Docx", 1) > 0 Then
        NM = Left(MyDOC, Len(MyDOC) - 5)
    Else
        NM = Left(MyDOC, Len(MyDOC) - 4)
    End If
    
    CN = ActiveDocument.InlineShapes.Count ' ȡ���ĵ���ͼƬ��
    TN = ActiveDocument.Tables.Count ' ȡ���ĵ��б����

    ' ��ԭ�ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
    MyDOC.Activate
    If CN > 0 Then
        For j = 1 To CN
            ActiveDocument.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
    End If

    ' ��ԭ�ĵ��еı���������(TAB1, TAB2, ...)���Ա㿽������
    If TN > 0 Then
        For j = 1 To TN
            ActiveDocument.Tables(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TAB" & j
        Next j
    End If

    '-----------------------------------------------------
    ' �½�һ���հ��ĵ�, ������ʱ�洢ͼƬ��ָ��ΪDOC_CN
    If CN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_CN = Application.ActiveWindow.Document
        
        ҳ������ 'ִ�С�ҳ�����á�������
        
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
        
        ' �����ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
        DOC_CN.Activate
        For j = 1 To CN
            DOC_CN.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
        
        DOC_CN.Activate
        'ActiveDocument.SaveAs FileName:=NM & "_ͼƬ.doc", FileFormat:=wdFormatDocument
        
    End If

    '-----------------------------------------------------
    ' �½�һ���հ��ĵ�, ������ʱ�洢���ָ��ΪDOC_TN
    If TN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_TN = Application.ActiveWindow.Document
        
        ҳ������ 'ִ�С�ҳ�����á�������
        
        ' �ѱ�񿽱����½��ĵ���
        MyDOC.Activate
        For j = 1 To TN
            MyDOC.Tables(j).Range.Copy
            DOC_TN.Activate
            Selection.EndKey Unit:=wdStory
            Selection.TypeParagraph
            Selection.Paste
            MyDOC.Activate
        Next j
        
        ' �����ĵ��еı���������(TAB1, TAB2, ...)���Ա㿽������
        DOC_TN.Activate
        For j = 1 To TN
            ActiveDocument.Tables(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TAB" & j
        Next j
        
        ' �����ĵ��еı��Ӧ�ú꣺���B
        DOC_TN.Activate
        For j = 1 To TN
            DOC_TN.Tables(j).Select
            ���B
        Next j
        
        DOC_TN.Activate
        'ActiveDocument.SaveAs FileName:=NM & "_���.doc", FileFormat:=wdFormatDocument
        
    End If
    
    '-----------------------------------------------------

    ' ɾ��ԭ�ĵ��еı��
    MyDOC.Activate
    For j = 1 To TN
        If MyDOC.Tables.Count > 0 Then
            MyDOC.Tables(1).Select
            Selection.Cut
        End If
    Next j
    
    ' ɾ��ԭ�ĵ��е�ͼƬ
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
    
    '�½��հ��ĵ�����ԭ�ĵ��е��ı����������ĵ�����ִ�С�A00_��ҳ��ʽ��������
    
    Documents.Add DocumentType:=wdNewBlankDocument
    Set DOC_OK = Application.ActiveWindow.Document
    
    ҳ������ 'ִ�С�ҳ�����á�������
    
    MyDOC.Activate
    MyText = MyDOC.Content.Text
    DOC_OK.Activate
    DOC_OK.Content.Text = MyText
    'A00_���ʾ�����Ϣ 'ִ�С�A00_���ʾ�����Ϣ��������
    A00_��ҳ��ʽ  'ִ�С�A00_��ҳ��ʽ��������
    'A00_���ʾ�����Ϣ 'ִ�С�A00_���ʾ�����Ϣ��������

    '-----------------------------------------------------
    
    '�ѱ�񿽱�����
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
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    
    '��ͼƬ��������
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
    
    �ո��滻
    A01_�����Ӵֱ��ϼ���
    A00_�����ӱ�ͷ��ע
    A00_�����ӱ�ע
    A00_��ע
    A00_PMI

    '-----------------------------------------------------

   Application.ScreenUpdating = True '�ָ���Ļ����


    DOC_OK.Activate

    '��׼������ĵ����,����_OK
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument
    ' ActiveDocument.SaveAs FileName:=NM & "_TB.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & "_OK.doc")

    'MsgBox "�ĵ��Ѿ���Ҫ���׼�������ĵ����У� " & Chr(13) & "    " & TN & " �ű��" & Chr(13) & "    " & CN & " ��ͼƬ" & Chr(13) & "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��"
    
End Sub

Sub A00_��ͼ��N()

    '�趨�ĵ�����Ŀ¼
    Dim FD1, FD2 As String
    
    FD1 = "C:\Users\zlzx-dhg\Desktop\00 OK_DOC"
    FD2 = "D:\00 OK_DOC"
    st = VBA.Timer '�������м�ʱ��

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

    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next '���Դ���

    '����ĵ����и���ʽͼƬ������ת��ΪǶ��ʽͼƬ
    If ActiveDocument.Shapes.Count > 0 Then
        For Each oShape In ActiveDocument.Shapes
            oShape.ConvertToInlineShape
        Next
    End If

    If ActiveDocument.Tables.Count > 0 Then
        A00_ȥ��ͷ��ע
    End If
    Selection.HomeKey Unit:=wdStory
    A00_ɾ����ҳ�ո�
    Selection.HomeKey Unit:=wdStory

    ' ��������������NM -- �ļ����� CN -- ͼƬ��  TN--�����
    Dim NM As String
    Dim CN As Integer
    Dim TN As Integer
        
    ' ȡ�õ�ǰ�ļ���
    Set MyDOC = Application.ActiveWindow.Document 'ָ��Ҫ������ĵ�ΪMyDoc
    If InStr(1, MyDOC, "Docx", 1) > 0 Then
        NM = Left(MyDOC, Len(MyDOC) - 5)
    Else
        NM = Left(MyDOC, Len(MyDOC) - 4)
    End If
    
    CN = ActiveDocument.InlineShapes.Count ' ȡ���ĵ���ͼƬ��
    TN = ActiveDocument.Tables.Count ' ȡ���ĵ��б����

    ' ��ԭ�ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
    MyDOC.Activate
    If CN > 0 Then
        For j = 1 To CN
            ActiveDocument.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
    End If

    ' ��ԭ�ĵ��еı���������(TAB1, TAB2, ...)���Ա㿽������
    If TN > 0 Then
        For j = 1 To TN
            ActiveDocument.Tables(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TAB" & j
        Next j
    End If

    '-----------------------------------------------------
    ' �½�һ���հ��ĵ�, ������ʱ�洢ͼƬ��ָ��ΪDOC_CN
    If CN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_CN = Application.ActiveWindow.Document
        
        ҳ������ 'ִ�С�ҳ�����á�������
        
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
        
        ' �����ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
        DOC_CN.Activate
        For j = 1 To CN
            DOC_CN.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
        
        DOC_CN.Activate
        'ActiveDocument.SaveAs FileName:=NM & "_ͼƬ.doc", FileFormat:=wdFormatDocument
        
    End If

    '-----------------------------------------------------
    ' �½�һ���հ��ĵ�, ������ʱ�洢���ָ��ΪDOC_TN
    If TN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_TN = Application.ActiveWindow.Document
        
        ҳ������ 'ִ�С�ҳ�����á�������
        
        ' �ѱ�񿽱����½��ĵ���
        MyDOC.Activate
        For j = 1 To TN
            MyDOC.Tables(j).Range.Copy
            DOC_TN.Activate
            Selection.EndKey Unit:=wdStory
            Selection.TypeParagraph
            Selection.Paste
            MyDOC.Activate
        Next j
        
        ' �����ĵ��еı���������(TAB1, TAB2, ...)���Ա㿽������
        DOC_TN.Activate
        For j = 1 To TN
            ActiveDocument.Tables(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TAB" & j
        Next j
        
        ' �����ĵ��еı��Ӧ�ú꣺���B
        DOC_TN.Activate
        For j = 1 To TN
            DOC_TN.Tables(j).Select
            ���B
        Next j
        
        DOC_TN.Activate
        'ActiveDocument.SaveAs FileName:=NM & "_���.doc", FileFormat:=wdFormatDocument
        
    End If
    
    '-----------------------------------------------------

    ' ɾ��ԭ�ĵ��еı��
    MyDOC.Activate
    For j = 1 To TN
        If MyDOC.Tables.Count > 0 Then
            MyDOC.Tables(1).Select
            Selection.Cut
        End If
    Next j
    
    ' ɾ��ԭ�ĵ��е�ͼƬ
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
    
    '�½��հ��ĵ�����ԭ�ĵ��е��ı����������ĵ�����ִ�С�A00_��ҳ��ʽ��������
    
    Documents.Add DocumentType:=wdNewBlankDocument
    Set DOC_OK = Application.ActiveWindow.Document
    
    ҳ������ 'ִ�С�ҳ�����á�������
    
    MyDOC.Activate
    MyText = MyDOC.Content.Text
    DOC_OK.Activate
    DOC_OK.Content.Text = MyText
    A00_��ҳ��ʽ  'ִ�С�A00_��ҳ��ʽ��������

    '-----------------------------------------------------
    
    '�ѱ�񿽱�����
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
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    
    '��ͼƬ��������
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
    
    �ո��滻
    A01_�����Ӵֱ��ϼ���
    A00_��ע
    A00_PMI
    A00_�����Ӵ־���
    A00_��λ����_����
    A00_��ע_����
    A00_�����_����

    '-----------------------------------------------------

   Application.ScreenUpdating = True '�ָ���Ļ����


    DOC_OK.Activate

    '��׼������ĵ����,����_OK
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument
    ' ActiveDocument.SaveAs FileName:=NM & "_TB.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & "_OK.doc")

    'MsgBox "�ĵ��Ѿ���Ҫ���׼�������ĵ����У� " & Chr(13) & "    " & TN & " �ű��" & Chr(13) & "    " & CN & " ��ͼƬ" & Chr(13) & "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��"
    
End Sub

Sub A00_��ͼ��_��Ϣ����()

    '�趨�ĵ�����Ŀ¼
    Dim FD1 As String
    Dim FD2 As String
    
    FD1 = "C:\Users\zlzx-dhg\Desktop\00 OK_DOC"
    FD2 = "D:\00 OK_DOC"
    st = VBA.Timer '�������м�ʱ��

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

    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next '���Դ���

    '����ĵ����и���ʽͼƬ������ת��ΪǶ��ʽͼƬ
    If ActiveDocument.Shapes.Count > 0 Then
        For Each oShape In ActiveDocument.Shapes
            oShape.ConvertToInlineShape
        Next
    End If

    If ActiveDocument.Tables.Count > 0 Then
        A00_ȥ��ͷ��ע
    End If
    Selection.HomeKey Unit:=wdStory
    A00_ɾ����ҳ�ո�
    Selection.HomeKey Unit:=wdStory

    ' ��������������NM -- �ļ����� CN -- ͼƬ��  TN--�����
    Dim NM As String
    Dim CN As Integer
    Dim TN As Integer
        
    ' ȡ�õ�ǰ�ļ���
    Set MyDOC = Application.ActiveWindow.Document 'ָ��Ҫ������ĵ�ΪMyDoc
    If InStr(1, MyDOC, "Docx", 1) > 0 Then
        NM = Left(MyDOC, Len(MyDOC) - 5)
    Else
        NM = Left(MyDOC, Len(MyDOC) - 4)
    End If
    
    CN = ActiveDocument.InlineShapes.Count ' ȡ���ĵ���ͼƬ��
    TN = ActiveDocument.Tables.Count ' ȡ���ĵ��б����

    ' ��ԭ�ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
    MyDOC.Activate
    If CN > 0 Then
        For j = 1 To CN
            ActiveDocument.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
    End If

    ' ��ԭ�ĵ��еı���������(TAB1, TAB2, ...)���Ա㿽������
    If TN > 0 Then
        For j = 1 To TN
            ActiveDocument.Tables(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TAB" & j
        Next j
    End If

    '-----------------------------------------------------
    ' �½�һ���հ��ĵ�, ������ʱ�洢ͼƬ��ָ��ΪDOC_CN
    If CN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_CN = Application.ActiveWindow.Document
        
        ҳ������ 'ִ�С�ҳ�����á�������
        
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
        
        ' �����ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
        DOC_CN.Activate
        For j = 1 To CN
            DOC_CN.InlineShapes(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TU" & j
        Next j
        
        DOC_CN.Activate
        'ActiveDocument.SaveAs FileName:=NM & "_ͼƬ.doc", FileFormat:=wdFormatDocument
        
    End If

    '-----------------------------------------------------
    ' �½�һ���հ��ĵ�, ������ʱ�洢���ָ��ΪDOC_TN
    If TN > 0 Then
        Documents.Add DocumentType:=wdNewBlankDocument
        Set DOC_TN = Application.ActiveWindow.Document
        
        ҳ������ 'ִ�С�ҳ�����á�������
        
        ' �ѱ�񿽱����½��ĵ���
        MyDOC.Activate
        For j = 1 To TN
            MyDOC.Tables(j).Range.Copy
            DOC_TN.Activate
            Selection.EndKey Unit:=wdStory
            Selection.TypeParagraph
            Selection.Paste
            MyDOC.Activate
        Next j
        
        ' �����ĵ��еı���������(TAB1, TAB2, ...)���Ա㿽������
        DOC_TN.Activate
        For j = 1 To TN
            ActiveDocument.Tables(j).Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
            Selection.InsertAfter Text:="TAB" & j
        Next j
        
        ' �����ĵ��еı��Ӧ�ú꣺���B
        DOC_TN.Activate
        For j = 1 To TN
            DOC_TN.Tables(j).Select
            ���B
        Next j
        
        DOC_TN.Activate
        'ActiveDocument.SaveAs FileName:=NM & "_���.doc", FileFormat:=wdFormatDocument
        
    End If
    
    '-----------------------------------------------------

    ' ɾ��ԭ�ĵ��еı��
    MyDOC.Activate
    For j = 1 To TN
        If MyDOC.Tables.Count > 0 Then
            MyDOC.Tables(1).Select
            Selection.Cut
        End If
    Next j
    
    ' ɾ��ԭ�ĵ��е�ͼƬ
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
    
    '�½��հ��ĵ�����ԭ�ĵ��е��ı����������ĵ�����ִ�С�A00_��ҳ��ʽ��������
    
    Documents.Add DocumentType:=wdNewBlankDocument
    Set DOC_OK = Application.ActiveWindow.Document
    
    ҳ������ 'ִ�С�ҳ�����á�������
    
    MyDOC.Activate
    MyText = MyDOC.Content.Text
    DOC_OK.Activate
    DOC_OK.Content.Text = MyText
    A00_��ҳ��ʽ  'ִ�С�A00_��ҳ��ʽ��������

    '-----------------------------------------------------
    
    '�ѱ�񿽱�����
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
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    
    '��ͼƬ��������
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
    
    �ո��滻
    A01_�����Ӵֱ��ϼ���
    A00_��ע
    A00_PMI
    A00_�����Ӵ־���
    A00_��λ_����_��Ϣ����
    A00_��ע_����

    '-----------------------------------------------------

   Application.ScreenUpdating = True '�ָ���Ļ����


    DOC_OK.Activate

    '��׼������ĵ����,����_OK
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument
    ' ActiveDocument.SaveAs FileName:=NM & "_TB.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & "_OK.doc")

    'MsgBox "�ĵ��Ѿ���Ҫ���׼�������ĵ����У� " & Chr(13) & "    " & TN & " �ű��" & Chr(13) & "    " & CN & " ��ͼƬ" & Chr(13) & "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��"
    
End Sub

Sub A00_���()
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
    
    '����һ�еĸ߶�����Ϊ1����
    Selection.SelectRow
    Selection.Rows.Height = CentimetersToPoints(1#)
    
    '���������Ϊ����
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Tables(1).Rows.Alignment = wdAlignRowCenter
    
    '�����ڵ������
    Selection.Tables(1).Select
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    
    '�̶������п�
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
    MsgBox "�ף�����㲻�ڱ���У�������ִ�б��������������" & Chr(13) & _
           "�뽫�����ŵ��������ⵥԪ���У�" & Chr(13) & _
           "Ȼ����ִ�б��꣬лл��"
   End If
      Application.ScreenUpdating = True '�ָ���Ļ����

End Sub

Sub A00_�������()
    Dim AA As Variant
    AA = Array("ëʢ��", "������", "���ǿ", "��־��", "������", "�йظ�����") ' ���巢��������
    
    '�ж��Ƿ����ִ�б�������������������б����С����ߡ�
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="���ߣ�"
        If .Found = False Then
            MsgBox "�ף���ǰ�ĵ�������ִ�б��������������" & Chr(13) _
            & "���ǵ����˰ɣ�������ȷ�����˳��ɣ�"
            Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
            Exit Sub
        End If
    End With
    
    On Error Resume Next
    Application.ScreenUpdating = False '�ر���Ļ����
    
   ' A00_��ҳ��ʽ
    
    For i = 0 To UBound(AA)
    
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
        For j = 1 To ActiveDocument.Paragraphs.Count
            With Selection.Find
                .ClearFormatting
                .Execute FindText:="����" & AA(i) & "��"
                If .Found = True Then
                   ' .Parent.Expand Unit:=wdParagraph
                    Selection.Range.Font.Bold = True
                End If
            End With
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        Next j
    Next i
        
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
        For j = 1 To ActiveDocument.Paragraphs.Count
            With Selection.Find
                .ClearFormatting
                .Execute FindText:="���ߣ�"
                If .Found = True Then
                    Selection.MoveRight Unit:=wdCharacter, Count:=1
                    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
                    Selection.Font.Bold = True
                    Selection.Paragraphs(1).Range.Select
                    Selection.Font.NameFarEast = "����"
                End If
            End With
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        Next j
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_���ļӴ�()

    '�����ض����ַ������磺"��"��Ȼ�󽫸��ַ���ǰ�����ݼӴ�
    
    Application.ScreenUpdating = False '�ر���Ļ����
    
    '���������Ŀո��滻Ϊһ�����Ŀո�
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(32) & Chr(32)
        .Replacement.Text = Chr(-24159)
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    '���������Ŀո��滻Ϊһ�����Ŀո�
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(32)
        .Replacement.Text = Chr(-24159)
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll


    '�ж��Ƿ����ִ�б�������������������б����С�����
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="��" & Chr(-24159)
        If .Found = False Then
            MsgBox "�ף���ǰ�ĵ�������ִ�б��������������" & Chr(13) _
            & "�����ã����ǵ����˰ɣ�������ȷ�����˳��ɣ�"
            Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
            Exit Sub
        End If
    End With

    '������ǰ�����ݼӴ�
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    For j = 0 To ActiveDocument.Paragraphs.Count - 1
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="��" & Chr(-24159)
        If .Found = True Then
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            Selection.Font.Bold = True
        End If
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next j
    
    '���½��д�������+����
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    For j = 0 To ActiveDocument.Paragraphs.Count - 1
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="��" & Chr(-24159)
        If .Found = True Then
            'Selection.MoveRight Unit:=wdCharacter, Count:=1
            'Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            Selection.Paragraphs(1).Range.Select
            Selection.Font.NameFarEast = "����"
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
    
    '�Խڽ��д�������Ӵ�+����
    Selection.WholeStory
    Selection.HomeKey Unit:=wdStory
    For j = 0 To ActiveDocument.Paragraphs.Count - 1
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="��" & Chr(-24159)
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
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_֪ͨ()
    
    A00_��ҳ��ʽ
    
    '�ж��Ƿ����ִ�б��������������Ҫ����֪ͨ����
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="��" & Chr(13)
        If .Found = True Then
            Selection.Paragraphs(1).Range.Select
            s = Selection.Paragraphs(1).Range.Text
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
        End If
    End With
    
    '�ĺž���
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=Chr(-24142)
        If .Found = True Then
            Selection.Paragraphs(1).Range.Select
            s = Selection.Paragraphs(1).Range.Text
            Selection.Paragraphs(1).Range.Text = Trim(s)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            'Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
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
    
    '�жϷ�������
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="��" & Chr(13)
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
            Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
        End If
    End With
    
End Sub

Sub A01_������ϵ��ʽ����׼��()
    ҳ������
    ActiveDocument.Tables(1).Select
    A01_������ϵ��ʽ����׼��1
    ActiveDocument.Tables(2).Select
    A01_������ϵ��ʽ����׼��2
    ActiveDocument.Tables(3).Select
    A01_������ϵ��ʽ����׼��3
End Sub
Sub A01_������ϵ��ʽ����׼��1()
    
    ���B
    
    Selection.SelectRow
    Selection.Font.Bold = True '��ͷ����Ӵ�
    
    Selection.Tables(1).Select '�����и���Сֵ0.8����
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.8)
    
    Selection.SelectColumn '��һ�о���
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Columns(2).Select '�ڶ��о���
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Columns(3).Select '�ڶ��о���
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Columns(4).Select '�ڶ��о���
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Select '���ñ����Ϊ1/4����
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
    'ȫ���������Ϊ10%��ɫ
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = wdColorGray10
    
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectRow '��һ�о���
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.MoveLeft Unit:=wdCharacter, Count:=1

End Sub

Sub A01_������ϵ��ʽ����׼��2()

    ���B
    
    Selection.SelectRow
    Selection.Font.Bold = True '��ͷ����Ӵ�
    
    Selection.Tables(1).Select '�����и���Сֵ0.8����
    Selection.Rows.HeightRule = wdRowHeightAtLeast
    Selection.Rows.Height = CentimetersToPoints(0.8)
    
    Selection.SelectColumn '��һ�о���
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Columns(2).Select '�ڶ��о���
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Columns(3).Select  '�����о���
    Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Selection.Font.Color = wdColorAutomatic
    
    Selection.Tables(1).Select '���ñ����Ϊ1/4����
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
    'ȫ���������Ϊ10%��ɫ
    Selection.Shading.Texture = wdTextureNone
    Selection.Shading.ForegroundPatternColor = wdColorAutomatic
    Selection.Shading.BackgroundPatternColor = wdColorGray10
    
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectRow '��һ�о���
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.MoveLeft Unit:=wdCharacter, Count:=1

End Sub

Sub A01_������ϵ��ʽ����׼��3()
    Application.Run MacroName:="Normal.NewMacros.���B"
    
    Selection.SelectRow
    Selection.Cells.Merge
    Selection.Font.Bold = True '��ͷ����Ӵ�
    
    Selection.Tables(1).Select '�����и���Сֵ0.8����
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

Sub A00_��ί����滻()
    Dim AA, BB As Variant
    Application.ScreenUpdating = False '�ر���Ļ����
    AA = Array("����ί", "���Ų�", "����ί", "ס����", "������", "���粿", "������", "��ͨ��")
    BB = Array("��չ�ĸ�ί", "��ҵ����Ϣ����", "��������ί", "ס�����罨�貿", "����������", "������Դ��ᱣ�ϲ�", "������Դ��", "��ͨ���䲿")
    For j = 0 To UBound(AA)
        Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = AA(j)
            .Replacement.Text = BB(j)
            .Wrap = wdFindStop
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next j
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_���淶�ĺ��滻()
    Application.ScreenUpdating = False '�ر���Ļ����
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    Do While True
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=Chr(91) '���ҡ�[��
        If .Found = True Then
            Selection.MoveRight Unit:=wdCharacter, Count:=5, Extend:=wdExtend
            Set MyRange = Selection.Range
            If Right(MyRange.Text, 1) = Chr(93) Then '�жϵ�5���ַ��Ƿ�Ϊ��]��
                Z1 = Right(Left(MyRange.Text, 2), 1)
                Z2 = Right(Left(MyRange.Text, 3), 1)
                Z3 = Right(Left(MyRange.Text, 4), 1)
                Z4 = Right(Left(MyRange.Text, 5), 1)
                Z5 = Right(MyRange.Text, 1)
                TT = Chr(-24142) & Z1 & Z2 & Z3 & Z4 & Chr(-24141)
                '�ж��м���ĸ��ַ��Ƿ�������,����ǣ����滻Ϊ�������ţ����򣬺���
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
    Application.ScreenUpdating = True '�ָ���Ļ����
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    
    
End Sub


Sub A0_OA���_�����ļ�()

N = ActiveDocument.Shapes.Count ' ȡ���ĵ���ͼƬ��

MsgBox N

    For Each s In ActiveDocument.Shapes
        With s.TextFrame
            If .HasText Then MsgBox .TextRange.Text
        End With
    Next

    Selection.MoveDown Unit:=wdLine, Count:=1
End Sub

Sub A00_��ע�����()

    '�����ض����ַ������磺"["��Ȼ�󽫸��ַ���ǰ�����ݼӴ�
    
    Application.ScreenUpdating = False '�ر���Ļ����
    
    '�ж��Ƿ����ִ�б�������������������б����С�[ ]��
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
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
            Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
            Exit Do
        End If
    End With
    Loop
    
    MsgBox i
    
    Application.ScreenUpdating = True '�ָ���Ļ����
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


Sub A00_ɾ�������()

    Application.ScreenUpdating = False '�ر���Ļ����
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
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub


Sub A01_��ת��()
    
    Application.ScreenUpdating = False '�ر���Ļ����
    Dim TN, RN As Integer
    Dim TT As String
    TT = Selection.Range.Text
    TN = 0
    
    If Len(TT) < 2 Then
        MsgBox "��ѡ����Ҫת�������ݣ�"
    Else
        RN = Selection.Paragraphs.Count
        If InStr(1, TT, Chr(9), 1) > 0 Then
            For Each st In Selection.Paragraphs(1).Range.Characters
                If st = Chr(9) Then
                    TN = TN + 1
                 End If
            Next st
            Selection.ConvertToTable Separator:=wdSeparateByTabs, NumColumns:=TN + 1, NumRows:=RN, AutoFitBehavior:=wdAutoFitFixed
            Application.Run MacroName:="���B"
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
    Application.ScreenUpdating = True '�ָ���Ļ����
    
End Sub


Function BackTXT(TT) '����ɾ��ָ���ַ�������ı�
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
Sub A00_��λ�о���()
    On Error Resume Next
    Set MyTab = Selection.Tables(1)
    L = MyTab.Columns.Count
    For i = 1 To MyTab.Columns.Count
        MyTab.Cell(1, i).Select
        TT = Selection.Range.Text
        If InStr(1, TT, "��", 1) > 0 And InStr(1, TT, "λ", 1) > 0 Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.SelectColumn
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i
    Selection.Tables(1).Cell(1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1

End Sub
Sub A01_����()  '���ܣ����Ժ�����

    Set MyTab = Selection.Tables(1)
    L = MyTab.Columns.Count
    For i = 1 To MyTab.Columns.Count
        MyTab.Cell(1, i).Select
        TT = Selection.Range.Text
        If InStr(1, TT, "��", 1) > 0 And InStr(1, TT, "λ", 1) > 0 Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.SelectColumn
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    Next i
    Selection.Tables(1).Cell(1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1

End Sub
Sub A00_ɾ���ض�����()
    'MsgBox Len(Selection.Paragraphs(1).Range.Text)
    
    
    On Error Resume Next
    Selection.HomeKey Unit:=wdStory '��λ���ĵ���ͷ
    For Each aPara In ActiveDocument.Paragraphs
        LN = Len(aPara.Range.Text)
        If LN = 4 And Left(aPara.Range.Text, 2) = "����" Then
            aPara.Range.Select
            Selection.Delete
            Selection.Delete
        End If
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Next aPara
    'Selection.Range.PasteSpecial DataType:=wdPasteText
    Selection.HomeKey Unit:=wdStory
End Sub
Sub A00_����������()
    
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
                MsgBox "��" & H1 & "�е�" & L1 & "�п�ʼΪ����"
               Exit For
            End If
    Next L
    If H1 > 1 And Val(TT) <> 0 Then Exit For
    Next H

End Sub

Sub A00_��һ�������������()
    
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
    'MsgBox "��󳤶�Ϊ��" & M & "  ��С����Ϊ��" & N & "  ���" & M - N
    
    If M - N < 4 Then
        Selection.SelectColumn
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End If
    MyTab.Cell(1, 1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub A00_����ߺ͵��߼Ӵ�()
    
    Dim TB As Table '���������
    
    Application.ScreenUpdating = False '�ر���Ļ����
    
    '���û��ѡ������ʾ�û�ѡ��Ҫ����ı��
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "��ѡ����"
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
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_���ÿ�ж��뷽ʽ()
    
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
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True '�ָ���Ļ����
    
End Sub

Sub A00_����������Ҷ���()
    
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
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
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
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_�ر������ĵ����˳�()
    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Application.Quit SaveChanges:=wdDoNotSaveChanges
End Sub

Sub A00_��ע()
    
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="������ע"
        If .Found = True Then
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            Selection.Font.Bold = True
            Set MyRange = ActiveDocument.Range(start:=Selection.Paragraphs(1).Range.start, End:=ActiveDocument.Range.End)
            MyRange.Font.NameFarEast = "����"
        Else
            Exit Sub
        End If
    End With

End Sub

Sub A00_PMI()
    
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="�����й�������ɹ����ϻ�"
        If .Found = True Then
            Selection.MoveRight Unit:=wdCharacter, Count:=1
            Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
            Selection.Font.NameFarEast = "����"
            Selection.Paragraphs(1).Range.Text = "�й�������ɹ����ϻ�" & Chr(13)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.TypeBackspace
            Selection.Paragraphs(1).Range.Select
            Selection.Font.NameFarEast = "����"
            Selection.Paragraphs(1).Range.Text = "����ͳ�ƾַ���ҵ��������" & Chr(13)
            Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Else
            Exit Sub
        End If
    End With
    Selection.HomeKey Unit:=wdStory

End Sub


Sub A00_��ǰ�ĵ�ת��Ϊ���ı�()
    
    ҳ������
    Selection.HomeKey Unit:=wdStory
    Selection.WholeStory
    TT = ActiveDocument.Content.Text
    Documents.Add DocumentType:=wdNewBlankDocument
    ҳ������
    ActiveDocument.Content.Text = TT
    
End Sub


Sub A00_���ʾ�����Ϣ()
    
    Application.ScreenUpdating = False '�ر���Ļ����
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .Text = "����Ŀ¼"
        .Replacement.Text = ""
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    ActiveDocument.Save
    A00_ȥ��ע
    A00_ȥ��ͷ
    A00_ȥ��ͷ
    A00_��ͼ��
    A00_IEI
    
    Application.ScreenUpdating = True '�ָ���Ļ����

End Sub

Sub A00_IEI_�������()
    TT = Selection.Paragraphs(1).Range.Text
    TT = Left(TT, Len(TT) - 1)
    JH = Right(TT, 1)
    If Len(TT) < 40 And Len(TT) > 2 And JH <> "��" Then
        Selection.Paragraphs(1).Range.Select
        Selection.Font.Bold = True
        Selection.Paragraphs(1).Range.Text = Trim(TT) & Chr(13)
        Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        'Selection.MoveRight Unit:=wdCharacter, Count:=1
    End If
End Sub

Sub A00_IEI_ָ���ַ�������()
    Dim A As Variant
    A = Array("���ʾ�������", "���ʾ�������", "�����г���̬", "����۵�")
    
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
    
    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next
    Set MyDOC = ActiveDocument
    TN = MyDOC.Tables.Count
    
    Select Case TN
    Case 0
        Set RNG = MyDOC.Range
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_�������
            A00_IEI_ָ���ַ�������
        Next j
    Case 1
        Set RNG = MyDOC.Range(start:=MyDOC.Paragraphs(1).Range.start, End:=MyDOC.Tables(1).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_�������
            A00_IEI_ָ���ַ�������
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(1).Range.End + 1, End:=MyDOC.Range.End)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_�������
            A00_IEI_ָ���ַ�������
        Next j
    
    Case 2
        Set RNG = MyDOC.Range(start:=MyDOC.Paragraphs(1).Range.start, End:=MyDOC.Tables(1).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_�������
            A00_IEI_ָ���ַ�������
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(1).Range.End + 1, End:=MyDOC.Tables(2).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_�������
            A00_IEI_ָ���ַ�������
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(2).Range.End + 1, End:=MyDOC.Range.End)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_�������
            A00_IEI_ָ���ַ�������
        Next j
        
    Case 2
        Set RNG = MyDOC.Range(start:=MyDOC.Paragraphs(1).Range.start, End:=MyDOC.Tables(1).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_�������
            A00_IEI_ָ���ַ�������
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(1).Range.End + 1, End:=MyDOC.Range.End)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_�������
            A00_IEI_ָ���ַ�������
        Next j
    
    Case 3
        Set RNG = MyDOC.Range(start:=MyDOC.Paragraphs(1).Range.start, End:=MyDOC.Tables(1).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_�������
            A00_IEI_ָ���ַ�������
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(1).Range.End + 1, End:=MyDOC.Tables(2).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_�������
            A00_IEI_ָ���ַ�������
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(2).Range.End + 1, End:=MyDOC.Tables(3).Range.start - 1)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_�������
            A00_IEI_ָ���ַ�������
        Next j
        
        Set RNG = MyDOC.Range(start:=MyDOC.Tables(3).Range.End + 1, End:=MyDOC.Range.End)
        For j = 1 To RNG.Paragraphs.Count
            RNG.Paragraphs(j).Range.Select
            A00_IEI_�������
            A00_IEI_ָ���ַ�������
        Next j
        
    End Select
    
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub
Sub A00_��ĵ�ת��Ϊ���ı�()
    
    Dim MyRange As Range  '����һ����Χ��Range������
    '���û��ѡ��Χ����ָ����ΧΪ�����ĵ�
    On Error Resume Next
    If Len(Selection.Range.Text) = 0 Then Selection.WholeStory
    Set MyRange = Selection.Range '�趨��Χ����Ϊѡ��ķ�Χ
    MyRange.Copy
    Documents.Add DocumentType:=wdNewBlankDocument
    ҳ������
    Selection.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:=wdInLine, DisplayAsIcon:=False
    A00_��ҳ��ʽ
        
End Sub

Sub A00_ȥ��ͷ()
    
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TN As Integer
    
    Application.ScreenUpdating = False '�ر���Ļ����
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
    Application.ScreenUpdating = True '�ָ���Ļ����

End Sub

Sub A00_ȥ��ע()
    
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TN As Integer
    
    Application.ScreenUpdating = False '�ر���Ļ����
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
    Application.ScreenUpdating = True '�ָ���Ļ����

End Sub

Sub A00_�����ӱ�ע()
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TN As Integer
    
    Application.ScreenUpdating = False '�ر���Ļ����
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
        If InStr(1, TT, "ע��", 1) > 0 Or InStr(1, TT, "������Դ��", 1) > 0 Then
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.TypeBackspace
            Selection.Paragraphs(1).Range.Select
            B01_ѡ���ı���Ϊ��ע
        End If
    Next i
   
    Application.ScreenUpdating = True '�ָ���Ļ����

End Sub

Sub A00_�����()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ChrW(8226)
        .Replacement.Text = "��"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub A00_��ҳ�ո��滻()
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
        A(0) = TT & Chr(13) & "�ǵ���"
    
    For i = 2 To 7
        ReDim Preserve A(i)
        TT = TB1.Cell(1, i + 1).Range.Text
        TT = Left(TT, Len(TT) - 2)
        A(i - 1) = TT & Chr(13) & "�ǵ���"
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
    Documents("ר��").Activate
    Set DC1 = ActiveDocument
    Documents("���ҡ�ʮ���塱ʱ���Ļ���չ�ĸ�滮��Ҫ").Activate
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
    
    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next
    
    Set TB1 = ActiveDocument.Tables(1)
    Set TB2 = ActiveDocument.Tables(2)
    ReDim A(1)
        TT = TB1.Cell(1, 2).Range.Text
        TT = Left(TT, Len(TT) - 2)
        A(0) = TT & Chr(13) & "�ǵ���"
    
    For i = 2 To 7
        ReDim Preserve A(i)
        TT = TB1.Cell(1, i + 1).Range.Text
        TT = Left(TT, Len(TT) - 2)
        A(i - 1) = TT & Chr(13) & "�ǵ���"
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
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_IEI_TB5()
    Dim A As Variant
    Dim B As Variant
    
    Application.ScreenUpdating = False '�ر���Ļ����
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
    
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_IEI_TB1()
    Dim A As Variant
    Dim B As Variant
    
    Application.ScreenUpdating = False '�ر���Ļ����
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
    
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_IEI_TB2()
    Dim A As Variant
    Dim B As Variant
    
    Application.ScreenUpdating = False '�ر���Ļ����
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
    
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_IEI_TB3()
    Dim A As Variant
    Dim B As Variant
    
    Application.ScreenUpdating = False '�ر���Ļ����
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
    
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_IEI_TB()
    
    Dim RNG As Range
    Dim A, B, C, D, FF, Z As Variant
    Z = Array("ȫ����Ҫ�������", "����ԭ�ͼ۸�", "���޵ĺ���ɢ����ָ��", "������Ҫ�������Ʊָ��", "ȫ����Ҫ���һ������")
    Dim DC1 As Document
    Dim TB1 As Table
    Dim TB2 As Table
    
    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next
    
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="���ʾ�����Ϣ"
        If .Found = False Then
            MsgBox "������ִ�б�������������"
            Exit Sub
        End If
    End With
    
    Set DC1 = ActiveDocument
    Selection.HomeKey Unit:=wdStory
    'A00_ȥ��ͷ
    Application.ScreenUpdating = False '�ر���Ļ����
    Selection.HomeKey Unit:=wdStory
        
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=Z(0) & Chr(13) 'ȫ����Ҫ�������
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
            .Execute FindText:=Z(1) & Chr(13) '����ԭ�ͼ۸�
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
            .Execute FindText:=Z(2) & Chr(13) '���޵ĺ���ɢ����ָ��
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
        .Execute FindText:=Z(3) & Chr(13)  '������Ҫ�������Ʊָ��
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
                Do While True
                    Selection.MoveDown Unit:=wdLine, Count:=1
                    If Selection.Information(wdWithInTable) = True Then
                        Exit Do
                    End If
                Loop
                Set TB1 = Selection.Tables(1)
                
                '����б�ע�У��򽫱�ע�ı���ֵ������BZ����ɾ����ע��
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
                
                '�����ͷ�����У���ϲ�
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
                
                '�����8�е���������10����ϲ��йص�Ԫ��
                RN = TB1.Rows.Count
                If RN > 10 Then
                    For i = 2 To 10
                        TT = Left(TB1.Cell(i, 8).Range.Text, Len(TB1.Cell(i, 8).Range.Text) - 2)
                        If TT <> "����" Then
                            TB1.Cell(i, 8).Range.Select
                            Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
                            Selection.Cells.Merge
                        End If
                    Next i
                End If
            
            '��������ݸ�ֵ������D
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
            .Execute FindText:=Z(4) & Chr(13) '"ȫ����Ҫ���һ������"
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
    
    
    A00_��ҳ��ʽ
    A00_IEI
    Application.ScreenUpdating = False '�ر���Ļ����
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
    
    ���B
    
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
    
    ���B
    
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
    
    ���B
    
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
    
    ���B
    
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
    
    ���B
    
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
    
    Application.ScreenUpdating = True '�ָ���Ļ����

End Sub


Sub A00_���ת��()

    If Selection.Information(wdWithInTable) = True Then
        R = ActiveDocument.Tables(1).Rows.Count
        C = ActiveDocument.Tables(1).Columns.Count
    Else
        MsgBox "����㲻�ڱ���У�"
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

Sub A00_IEI_TB4_�ϲ��йص�Ԫ��()
    
    Dim TB1 As Table
    Dim RNG As Range
    
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="������Ҫ�������Ʊָ��"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
            Selection.MoveDown Unit:=wdLine, Count:=3
            If Selection.Information(wdWithInTable) = True Then
                Set TB1 = Selection.Tables(1)
                
                '����б�ע�У��򽫱�ע�ı���ֵ������BZ����ɾ����ע��
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
                
                '�����ͷ�����У���ϲ�
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
                
                '�����8�е���������10����ϲ��йص�Ԫ��
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
                        If TT <> "����" Then
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
Sub A00_����()
    A00_IEI_TB
    'A00_����������Ҷ���
    
End Sub
Sub A02_����()  '���ܣ����Ժ�����

End Sub

Sub A00_ȥ��ͷ��ע()
    
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim H As Integer
    Dim TT As String
    Dim DW As String
    Dim BZ As String
    
    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next
    Set MyDOC = ActiveDocument
    TN = MyDOC.Tables.Count
    
    Selection.HomeKey Unit:=wdStory
    For i = 1 To TN
        Set TB = MyDOC.Tables(i)
        'ȥ��ͷ
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
        
        'ȥ��ע
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
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_�ӱ�ͷ��ע()
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim H As Integer
    Dim TT As String
    Dim DW As String
    Dim BZ As String
    
    'Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next
    Set MyDOC = ActiveDocument
    Set TB = MyDOC.Tables(1)
    
    TB.Select
    
    '�������ͷ����ķǿ���
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
    
    '������ݰ�������λ��,�������һ���ǿ��У�ѡ���Ϊ����λ�����б�ͷ
    If InStr(1, TT, "��λ", 1) > 0 Then
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
        B01_ѡ�����ֱ�Ϊ��ͷ
    Else
         TXT = Left(TT, Len(TT) - 1)
         'MsgBox Right(TXT, 1)
         If Right(TXT, 1) = "��" Or Right(TXT, 1) = "��" Or Len(TXT) > 30 Then
            TB.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
        Else
            Selection.Paragraphs(1).Range.Select
            B01_ѡ�����ֱ�Ϊ��ͷ
        End If
    End If
    
    TB.Select
    
    '������������ķǿ���
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
    
    If InStr(1, TT, "ע��", 1) > 0 Or InStr(1, TT, "������Դ��", 1) > 0 Then
        Selection.Paragraphs(1).Range.Select
        Selection.Paragraphs(1).Range.Text = Trim(TT)
        Selection.Paragraphs(1).Range.Select
        B01_ѡ���ı���Ϊ��ע
    Else
        Selection.Paragraphs(1).Range.Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.TypeParagraph
    End If
    
    End Sub
Sub A00_�����ӱ�ͷ��ע()
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim H As Integer
    Dim TT As String
    Dim DW As String
    Dim BZ As String
    
    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next
    Set MyDOC = ActiveDocument
    'Set TB = MyDOC.Tables(1)
    TN = MyDOC.Tables.Count
    If TN > 0 Then
    For i = 1 To TN
    Set TB = MyDOC.Tables(i)
    
    TB.Select
    
    '�������ͷ����ķǿ���
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
    
    '������ݰ�������λ��,�������һ���ǿ��У�ѡ���Ϊ����λ�����б�ͷ
    If InStr(1, TT, "��λ", 1) > 0 Then
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
        B01_ѡ�����ֱ�Ϊ��ͷ
    Else
         TXT = Left(TT, Len(TT) - 1)
         'MsgBox Right(TXT, 1)
         If Right(TXT, 1) = "��" Or Right(TXT, 1) = "��" Or Len(TXT) > 30 Then
            TB.Select
            Selection.MoveLeft Unit:=wdCharacter, Count:=2
            Selection.TypeParagraph
        Else
            Selection.Paragraphs(1).Range.Select
            B01_ѡ�����ֱ�Ϊ��ͷ
        End If
    End If
    
    TB.Select
    
    '������������ķǿ���
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
    
    If InStr(1, TT, "ע��", 1) > 0 Or InStr(1, TT, "������Դ��", 1) > 0 Then
        Selection.Paragraphs(1).Range.Select
        Selection.Paragraphs(1).Range.Text = Trim(TT)
        Selection.Paragraphs(1).Range.Select
        B01_ѡ���ı���Ϊ��ע
    Else
        Selection.Paragraphs(1).Range.Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.TypeParagraph
    End If
    
    Next i
    End If
    
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub
Sub A00_ɾ������()
    If ActiveDocument.Tables.Count > 0 Then
        MsgBox "�ĵ����б�񣬲���ִ�б������"
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
Sub A00_ɾ��ѡ���������()
    
    Dim RNG As Range
    Set RNG = Selection.Range
    If Len(Trim(RNG.Text)) < 2 Then
        MsgBox "�ף���û��ѡ���ı�������ִ�б������"
        Exit Sub
    Else
        If RNG.Tables.Count > 0 Then
            MsgBox "�ĵ����б�񣬲���ִ�б������"
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




Sub A00_����N()
    
    A00_ɾ������
    'A00_ȥ��ͷ��ע
    'MsgBox "��ȥ��ͷ��ע"
    'A00_�����ӱ�ͷ��ע
    'MsgBox "�Ѽӱ�ͷ��ע"
End Sub


Sub A00_ɾ����ҳ�ո�()
    
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

Sub A00_�����Ӵ־���()
    
    Dim N As Integer
    Dim TT As String
    
    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next
    Selection.HomeKey Unit:=wdStory
    N = ActiveDocument.Tables.Count
    If N > 0 Then
        TT = "��" & N
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=TT
            If .Found = True Then
                Selection.HomeKey Unit:=wdStory
                For i = 1 To N
                    With Selection.Find
                        .ClearFormatting
                        .Execute FindText:="��" & i
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
    Application.ScreenUpdating = True '�ָ���Ļ����

End Sub

Sub A00_�����_����()
    Dim RNG As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim TT As String
    
    Application.ScreenUpdating = False '�ر���Ļ����
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
            If Right(TT, 1) <> "��" Or Right(TT, 1) <> "��" And Len(TT) < 40 Then
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
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_��λ����()
    
    TT = Selection.Range.Text
    TT = Trim(TT)
    TT = Left(TT, Len(TT) - 1)
    Set MyRange = Selection.Range
    If InStr(1, TT, "��λ", 1) > 0 And Selection.Paragraphs.Count = 1 Then
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

Sub A00_��λ����_����()
    Dim RNG1 As Range
    Dim RNG2 As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim TT As String
    
    Application.ScreenUpdating = False '�ر���Ļ����
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
            '������ݰ�������λ��,��ִ�С�A00_��λ���ҡ�������
            If InStr(1, TT, "��λ", 1) > 0 Then
                RNG1.Select
                Selection.Delete
                RNG2.Select
                A00_��λ����
            End If
            TB.Select
            Selection.MoveDown Unit:=wdLine, Count:=1
        Next i
    End If
    
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub
Sub A00_��λ_����_��Ϣ����()
    Dim RNG1 As Range
    Dim RNG2 As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim TT As String
    
    Application.ScreenUpdating = False '�ر���Ļ����
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
            '������ݰ�������λ��,��ִ�С�A00_��λ���ҡ�������
            If InStr(1, TT, "��λ", 1) > 0 Then
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
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_��ע_����()
    Dim RNG1 As Range
    Dim RNG2 As Range
    Dim MyDOC As Document
    Dim TB As Table
    Dim TN As Integer
    Dim TT As String
    
    Application.ScreenUpdating = False '�ر���Ļ����
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
            If InStr(1, TT, "ע��", 1) > 0 Or InStr(1, TT, "������Դ��", 1) > 0 Then
                RNG1.Select
                Selection.Delete
                RNG2.Select
                B01_ѡ���ı���Ϊ��ע
            End If
            TB.Select
            Selection.MoveDown Unit:=wdLine, Count:=1
        Next i
    End If
    
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub A00_�趨�ĵ�����Ŀ¼()

    '�趨�ĵ�����Ŀ¼
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

Sub A00_ʧ����ҵ()

    '��������
    Dim DOC1, DOC2 As Document
    Dim FD1, NM1, TT As String
    Dim AA() As Variant
    Dim RNG As Range
    Dim TB, TB2 As Table
    Dim i, RN, CN As Integer
    
    FD1 = "F:\03 SXQY"
    NM1 = "ͳ��������ʧ����ҵ��Ϣ��ʾ��"
    
    '�趨�ĵ�����Ŀ¼
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(FD1) = True Then
        ChangeFileOpenDirectory FD1
    Else
        Set A = fs.CreateFolder(FD1)
        ChangeFileOpenDirectory FD1
    End If
    
    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next
    
    Documents("ʧ����ҵ��ʾ��Ϣ.Docx").Activate
    Set DOC1 = ActiveDocument
    Documents("ʧ����ҵ��Ϣ��ʾģ��.Docx").Activate
    Set DOC2 = ActiveDocument
    'MsgBox ActiveDocument.Paragraphs(1).Range.Text
    DOC1.Activate
    'MsgBox ActiveDocument.Paragraphs(1).Range.Text
    Set TB = DOC1.Tables(1)
    RN = TB.Rows.Count '����
    CN = TB.Columns.Count '����
    'MsgBox "������" & RN & "������" & CN
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
    ҳ������
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
    Application.ScreenUpdating = True '�ָ���Ļ����
    
End Sub

Sub A00_OA���1()
    
    Dim TT As String
    Dim P1 As Integer
    Dim C As Variant
    C = Array(Chr(13), Chr(32), Chr(-24159), Chr(11))
    
    '���û��ѡ���ı�����ʾ�û�ѡ��
    If Len(Selection.Range.Text) = 0 Then
        MsgBox "��ע�⡿û��ѡ����Ϊ����ĵ��ļ������ı�" & Chr(13) & _
           "��ѡ�������ļ�������ѡ�У� " & Chr(13) & _
           "Ȼ����ִ�б��꣬лл��"
        Exit Sub
    End If
        
    TT = Selection.Range.Text
    
    'ȥ���ı���ָ���ַ���
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

Sub A00_OA���()

    Dim TT, FD1, FD2 As String
    Dim P1 As Integer
    Dim C As Variant
    
    '�趨�ĵ�����Ŀ¼
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

    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next '���Դ���
    
    '���û��ѡ���ı�����ʾ�û�ѡ��
    If Len(Selection.Range.Text) = 0 Then
        If Left(ActiveDocument.Paragraphs(1).Range.Text, Len(ActiveDocument.Paragraphs(1).Range.Text) - 1) = "�����������" Then NM1 = "�����������"
        If Left(ActiveDocument.Paragraphs(4).Range.Text, Len(ActiveDocument.Paragraphs(4).Range.Text) - 1) = "ÿ�յ���" Then NM1 = "ÿ�յ���"
        NM2 = ActiveDocument.Shapes(1).TextFrame.TextRange.Text
        NM2 = Left(NM2, Len(NM2) - 1)
        TT = NM1 & NM2
        If NM1 = "" Then
            MsgBox "��ע�⡿û��ѡ����Ϊ����ĵ��ļ������ı�" & Chr(13) & _
               "��ѡ�������ļ�������ѡ�У� " & Chr(13) & _
               "Ȼ����ִ�б��꣬лл��"
            Exit Sub
        End If
    Else
        TT = Selection.Range.Text
    End If
    
    'ȥ���ı���ָ���ַ���
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

