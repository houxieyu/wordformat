Attribute VB_Name = "NewMacros"


Sub ȫ��ת��()

    '��������ת���й����ݣ���س�ת��ΪӲ�س���ȫ�����֡���ĸת��Ϊ������֡���ĸ��
    Selection.WholeStory
    Dim C As Variant
    Dim D As Variant
    C = Array("^l", ",", ";", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", _
        "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", _
        "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", _
        "��", "��", "��", "��", "��", "��", "��")
    D = Array("^p", "��", "��", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
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



Sub ȫ������ת��S()
'
' ȫ������ת��S Macro
' ���� 2003-6-27 �� DHG ¼��
'
    Selection.WholeStory
    Dim A As Variant
    Dim B As Variant
    A = Array("��", "��", "��", "��", "��", "��", "��", "��", "��", "��")
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
Sub �������C()
'
' �������C Macro
' ���� 2003-6-27 �� DHG ¼��
'
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
End Sub
Sub �����ļ��еı�񲢸�ʽ��T()

   Application.ScreenUpdating = False '�ر���Ļ����

'
' �����ļ��еı�񲢸�ʽ��T Macro
' ���� 2003-6-27 �� DHG ¼��
'
' ��������������NM -- ��ǰ�ĵ����ļ����� T -- �ĵ��ı������
    Dim NM As String
    Dim T As Integer
    
' ȡ�õ�ǰ�ļ���
    Set MyDOC = Application.ActiveWindow.Document
'    MsgBox myDoc
    NM = Left(MyDOC, Len(MyDOC) - 4)
'    MsgBox NM
' �½�һ���հ��ĵ�
    Documents.Add DocumentType:=wdNewBlankDocument
    Set MyDocN = Application.ActiveWindow.Document

' �ѱ�񿽱����½��ĵ���
    MyDOC.Activate
    T = ActiveDocument.Tables.Count ' ȡ���ĵ��б�������
'    MsgBox "����ļ����� " & T & " �ű��"
For j = 1 To T
    ActiveDocument.Tables(j).Range.Copy
    MyDocN.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.Paste
    MyDOC.Activate
Next j

' �����ĵ��еı��Ӧ�ú꣺���B
    MyDocN.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
    ActiveDocument.Tables(j).Select
    Application.Run MacroName:="Normal.NewMacros.���B"
Next j
' ��ԭ�ĵ��еı���������(Tab1, Tab2, ...)���Ա㿽������
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
Next j

' ɾ��ԭ�ĵ��еı��
For j = 1 To T
    If ActiveDocument.Tables.Count > 0 Then
    ActiveDocument.Tables(1).Range.Cut
    End If
Next j

'���� ��׼��ҳ��ʽW Ӧ�õ��ĵ���
Selection.WholeStory
Selection.Copy
Documents.Add DocumentType:=wdNewBlankDocument
Selection.Range.PasteSpecial DataType:=wdPasteText

    Application.Run MacroName:="Normal.NewMacros.��׼��ҳ��ʽW"
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
    
       Application.ScreenUpdating = True '�ָ���Ļ����

    
End Sub
Sub ɾ�ĵ�ͼƬ�ͱ��Ӧ�ñ�׼��ҳ��ʽ()

   Application.ScreenUpdating = False '�ر���Ļ����

'
' ��������������NM -- �ļ����� N -- ͼƬ��  T--�����
    Dim NM As String
    Dim N As Integer
    Dim T As Integer
        
' ȡ�õ�ǰ�ļ���
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    N = ActiveDocument.InlineShapes.Count ' ȡ���ĵ���ͼƬ��
    T = ActiveDocument.Tables.Count ' ȡ���ĵ��б����

' ��ԭ�ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
    MyDOC.Activate
For j = 1 To N
    ActiveDocument.InlineShapes(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TU" & j
Next j
' ɾ��ԭ�ĵ��е�ͼƬ
For j = 1 To N
    If ActiveDocument.InlineShapes.Count > 0 Then
    ActiveDocument.InlineShapes(1).Range.Cut
    End If
Next j

' �½�һ���հ��ĵ�
    Documents.Add DocumentType:=wdNewBlankDocument
    Set MyDocN = Application.ActiveWindow.Document

' �ѱ�񿽱����½��ĵ���
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Range.Copy
    MyDocN.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.Paste
    MyDOC.Activate
Next j

' �����ĵ��еı��Ӧ�ú꣺���B
    MyDocN.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
    ActiveDocument.Tables(j).Select
    Application.Run MacroName:="Normal.NewMacros.���B"
Next j
' ��ԭ�ĵ��еı���������(Tab1, Tab2, ...)���Ա㿽������
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
Next j

' ɾ��ԭ�ĵ��еı��
For j = 1 To T
    If ActiveDocument.Tables.Count > 0 Then
    ActiveDocument.Tables(1).Range.Cut
    End If
Next j

'���� A00_��ҳ��ʽ Ӧ�õ��ĵ���
    A00_��ҳ��ʽ
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
   
   Application.ScreenUpdating = True '�ָ���Ļ����

'��׼������ĵ����,����_OK
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & ".doc")
    ActiveDocument.SaveAs FileName:=NM & ".htm", FileFormat:=wdFormatHTML
    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & ".doc")
    Documents.Open (NM & "_OK.doc")
    MsgBox "�ĵ��Ѿ���Ҫ���׼�������ĵ����У� " & Chr(13) _
     & "    " & T & " �ű��" & Chr(13) _
     & "    " & N & " ��ͼƬ" & Chr(13) _
     & "ͼƬ������ " & NM & ".files Ŀ¼��"
    

    
    
End Sub
Sub �����ĵ�ͼƬ�ͱ��һ�����ĵ�()
   Application.ScreenUpdating = False '�ر���Ļ����

'
' ��������������NM -- �ļ����� N -- ͼƬ��  T--�����
    Dim NM As String
    Dim N As Integer
    Dim T As Integer
    
    On Error Resume Next
        
' ȡ�õ�ǰ�ļ���
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    N = ActiveDocument.InlineShapes.Count ' ȡ���ĵ���ͼƬ��
    T = ActiveDocument.Tables.Count ' ȡ���ĵ��б����

' ��ԭ�ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
    MyDOC.Activate
For j = 1 To N
    ActiveDocument.InlineShapes(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TU" & j
Next j

' �½�һ���հ��ĵ�
    Documents.Add DocumentType:=wdNewBlankDocument
    Set MyDocN = Application.ActiveWindow.Document

' �ѱ�񿽱����½��ĵ���
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Range.Copy
    MyDocN.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.Paste
    MyDOC.Activate
Next j

' ɾ��ԭ�ĵ��е�ͼƬ
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

' ��ԭ�ĵ��еı���������(Tab1, Tab2, ...)���Ա㿽������
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
Next j

' �����ĵ��еı���������(Tab1, Tab2, ...)���Ա㿽������
    MyDocN.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
Next j

' �����ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
    MyDocN.Activate
For j = 1 To N
    ActiveDocument.InlineShapes(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TU" & j
Next j


' ɾ��ԭ�ĵ��еı��
    MyDOC.Activate
For j = 1 To T
    If ActiveDocument.Tables.Count > 0 Then
    ActiveDocument.Tables(1).Range.Cut
    End If
Next j

   Application.ScreenUpdating = True '�ָ���Ļ����


'��׼������ĵ����,����_OK
    ActiveDocument.SaveAs FileName:=NM & "_WZ.doc", FileFormat:=wdFormatDocument
    MyDocN.Activate
    ActiveDocument.SaveAs FileName:=NM & "_TB.doc", FileFormat:=wdFormatDocument
    MsgBox "�ĵ��Ѿ���Ҫ���׼�������ĵ����У� " & Chr(13) _
     & "    " & T & " �ű��" & Chr(13) _
     & "    " & N & " ��ͼƬ" & Chr(13)
    
End Sub

Sub A01_��ͼ��()

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
    NM = Left(MyDOC, Len(MyDOC) - 5)
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
        ActiveDocument.SaveAs FileName:=NM & "_ͼƬ.doc", FileFormat:=wdFormatDocument
        
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
        ActiveDocument.SaveAs FileName:=NM & "_���.doc", FileFormat:=wdFormatDocument
        
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
    ActiveDocument.SaveAs FileName:=NM & "_WZ.doc", FileFormat:=wdFormatDocument
    
    '-----------------------------------------------------
    
    '�½��հ��ĵ�����ԭ�ĵ��е��ı����������ĵ�����ִ�С�A00_��ҳ��ʽ��������
    
    Documents.Add DocumentType:=wdNewBlankDocument
    Set DOC_OK = Application.ActiveWindow.Document
    
    ҳ������ 'ִ�С�ҳ�����á�������
    
    MyDOC.Activate
    Selection.WholeStory
    Selection.Copy
    
    DOC_OK.Activate
    
    CommandBars("Office Clipboard").Visible = False
    Selection.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:= _
        wdInLine, DisplayAsIcon:=False
    
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
            .Execute FindText:="����TU" & j
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
    'A01_���������ͷ

    '-----------------------------------------------------

   Application.ScreenUpdating = True '�ָ���Ļ����


    DOC_OK.Activate

    '��׼������ĵ����,����_OK
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument
    ' ActiveDocument.SaveAs FileName:=NM & "_TB.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & "_OK.doc")

    MsgBox "�ĵ��Ѿ���Ҫ���׼�������ĵ����У� " & Chr(13) _
     & "    " & TN & " �ű��" & Chr(13) _
     & "    " & CN & " ��ͼƬ" & Chr(13) _
     & "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��"
    
End Sub
Sub A01_��ͼ��E()

' ��������������NM -- �ļ����� n -- ͼƬ��  t--�����
    Dim NM As String
    Dim N As Integer
    Dim T As Integer

    st = VBA.Timer '�������м�ʱ��

   Application.ScreenUpdating = False '�ر���Ļ����
   On Error Resume Next '���Դ���
   
   A01_���ϲ���ͷ��עE
        
' ȡ�õ�ǰ�ļ���
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    N = ActiveDocument.InlineShapes.Count ' ȡ���ĵ���ͼƬ��
    T = ActiveDocument.Tables.Count ' ȡ���ĵ��б����

' ��ԭ�ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
    MyDOC.Activate
For j = 1 To N
    ActiveDocument.InlineShapes(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TU" & j
Next j

' �½�һ���հ��ĵ�
    Documents.Add DocumentType:=wdNewBlankDocument
    Set MyDocN = Application.ActiveWindow.Document
    ��ҳҳ��

' �ѱ�񿽱����½��ĵ���
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Range.Copy
    MyDocN.Activate
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.Paste
    MyDOC.Activate
Next j

' ��ԭ�ĵ��е�ͼƬ���������ĵ�
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

' ��ԭ�ĵ��еı���������(Tab1, Tab2, ...)���Ա㿽������
    MyDOC.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
Next j

' �����ĵ��еı���������(Tab1, Tab2, ...)���Ա㿽������
    MyDocN.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
Next j

' �����ĵ��еı��Ӧ�ú꣺���E
    MyDocN.Activate
For j = 1 To T
    ActiveDocument.Tables(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TAB" & j
    ActiveDocument.Tables(j).Select
    ���E
Next j

' �����ĵ��е�ͼƬ��������(TU1, TU2, ...)���Ա㿽������
    MyDocN.Activate
For j = 1 To N
    ActiveDocument.InlineShapes(j).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeParagraph
    Selection.InsertAfter Text:="TU" & j
Next j


' ɾ��ԭ�ĵ��еı��
    MyDOC.Activate
For j = 1 To T
    If ActiveDocument.Tables.Count > 0 Then
    ActiveDocument.Tables(1).Range.Cut
    End If
Next j

'���� A01_Ӣ����ҳ��ʽ Ӧ�õ��ĵ���
    MyDOC.Activate
    
    ��ҳҳ��
    A01_Ӣ����ҳ��ʽ
    
    '�ѱ�񿽱�����
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
    
    '��ͼƬ��������
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
    �ո��滻
    A01_�����Ӵֱ��ϼ���E
    A01_�����Ӵֱ���е��ض���E
    A01_���������ͷE
    
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
    
    A01_Ӣ��ע�͸�ʽ

   Application.ScreenUpdating = True '�ָ���Ļ����


'��׼������ĵ����,����_OK
    ActiveDocument.SaveAs FileName:=NM & "_OK.doc", FileFormat:=wdFormatDocument
    MyDocN.Activate
   ' ActiveDocument.SaveAs FileName:=NM & "_TB.doc", FileFormat:=wdFormatDocument

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Open (NM & "_OK.doc")

    MsgBox "�ĵ��Ѿ���Ҫ���׼�������ĵ����У� " & Chr(13) _
     & "    " & T & " �ű��" & Chr(13) _
     & "    " & N & " ��ͼƬ" & Chr(13) _
     & "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��"
    
End Sub

Sub A01_Ӣ��ע�͸�ʽ()
    
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

Sub A01_����Ӵ�E()
    
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

Sub A01_Ӣ����ҳ��ʽ()
    
    'st = VBA.Timer '�������м�ʱ��
    Application.ScreenUpdating = False '�ر���Ļ����
    Selection.HomeKey Unit:=wdStory

    Selection.WholeStory
    Selection.Cut
    Selection.Collapse Direction:=wdCollapseStart
    Selection.Range.PasteSpecial DataType:=wdPasteText
    
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
    
    Selection.WholeStory

    Dim O As Variant
    Dim R As Variant
    O = Array("^l", "  ", "^p^p", "^p", "����^p")
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
    A01_����Ӵ�E
    �ӿ���
    
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p^p^p"
        .Replacement.Text = "^p^p"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    
    Selection.HomeKey Unit:=wdStory
        
    Application.ScreenUpdating = True '�ر���Ļ����
   ' MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��" '��ʾ�������е�ʱ��


End Sub

Sub �����ֵ������D()
'
' ALT+T
' �����ֵ������D Macro
' ���� 2003-5-23 �� DHG ¼��
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

Sub �����ĵ�����()

' ��������������NM -- �ļ����� N -- ͼƬ��  T--����� Z--����
    Dim NM As String
    Dim N As Integer
    Dim T As Integer
    Dim Z As Integer
    Dim doc As Document
    Dim docFound As Boolean

    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)  ' ȡ�õ�ǰ�ļ���
    N = ActiveDocument.InlineShapes.Count ' ȡ���ĵ���ͼƬ��
    T = ActiveDocument.Tables.Count ' ȡ���ĵ��б����
    Z = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharacters) '����
    D = Date
    MyDOC.Close SaveChanges:=wdDoNotSaveChanges

    For Each doc In Documents
        If InStr(1, doc.name, "���ͳ��.doc", 1) Then
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
    .LookIn = "D:\02 MyDOC\02 ͳ���ƶ�\ԭʼ�ļ�"
    .FileName = "*.DOC"
    If .Execute(SortBy:=msoSortByFileName, _
    SortOrder:=msoSortOrderAscending) > 0 Then
        For i = 1 To .FoundFiles.Count
    Documents.Open FileName:=.FoundFiles(i)
    Application.Run MacroName:="Normal.NewMacros.�����ĵ�����"

        Next i
    Else
        MsgBox "û�ҵ������ĵ�"
    End If
End With
 
End Sub

Sub ShowFolderList()
    Dim fs, f, f1, s, sf
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder("F:\��֤����")
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

Sub ��ͳ��()
    Selection.TypeText Text:="����ͳ�ƾְ칫��"
    Selection.TypeParagraph
    Selection.InsertDateTime DateTimeFormat:="EEEE��O��A��", InsertAsField:=False
    Selection.TypeParagraph
End Sub
Sub ����1()
    Selection.InsertDateTime DateTimeFormat:="yyyy'��'M'��'d'��'", InsertAsField:=True
End Sub
Sub ����2()
    Selection.InsertDateTime DateTimeFormat:="EEEE��O��A��", InsertAsField:=False
End Sub

Sub �س��滻()
    Application.ScreenUpdating = False '�ر���Ļ����

    '��ѡ����ı�������в���,���û��ѡ��,���Զ�ѡ�������ĵ�
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
    Application.ScreenUpdating = True '�ر���Ļ����

End Sub

Sub �����滻K()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "��"
        .Replacement.Text = "("
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "��"
        .Replacement.Text = ")"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub

Sub ���ߵ�λ()
    Selection.TypeText Text:="���������ߵ�λ����"
    
End Sub

Sub ��������()
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
        .Replacement.Text = "��"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "��"
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

Sub Ӣ�ı�M01()
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

Sub Ӣ�ı�M02()
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

Sub Ӣ�ı�M03()
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

Sub Ӣ�ı�M04()
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


Sub ͳ���ƶȱ�ͷ()

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

    Application.Run MacroName:="�س��滻"

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

Sub ͳ���ƶȱ��ת��()
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
    Selection.Style = ActiveDocument.Styles("����")
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
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 9
    End With
    
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="�ۺϻ������ƣ�"
    End With
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeParagraph
    
    Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="������λ"
                If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
            str = "������λ"
        Else
        .Execute FindText:="��Ч����"
            .Parent.Expand Unit:=wdParagraph
            str = "��Ч����"
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
    Selection.TypeText Text:="����������"
    Selection.TypeParagraph
    Selection.MoveDown Unit:=wdLine, Count:=2
    Selection.Paragraphs(1).Range.Delete
        Selection.WholeStory
    With Selection.Find
        .ClearFormatting
        .Execute FindText:="˵��"
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
    Application.Run MacroName:="�����"

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
    Application.Run MacroName:="�س��滻"

Selection.Tables(1).Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveLeft Unit:=wdCharacter, Count:=1

    Selection.InsertRowsBelow 1
    Selection.SelectRow
    Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
    
    Selection.TypeText Text:="��λ�����ˣ�"
        Selection.SelectCell
    With Selection.ParagraphFormat
        .SpaceAfter = 3
        .Alignment = wdAlignParagraphJustify
    End With

    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="����ˣ�"
           Selection.SelectCell
    With Selection.ParagraphFormat
        .SpaceAfter = 3
        .Alignment = wdAlignParagraphJustify
    End With
    
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="�������ڣ����������ꡡ�¡���"
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
        .NameFarEast = "����"
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

    Application.Run MacroName:="�س��滻"

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

    Application.Run MacroName:="�س��滻"



End Sub

Sub ɾ��Ӣ����ĸ()

    Dim C As Variant
    Dim D As Variant
    C = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", _
              "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", " ", "��")
    For i = 0 To 53
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = ""
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    

End Sub

Sub ɾ������()

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

Sub ʡ�ݸ�ʽ()

    Dim C As Variant
    Dim D As Variant
    C = Array("ȫ��", "����", "���", "�ӱ�", "ɽ��", "����", "����", "�Ϻ�", "����", "�㽭", "����", "����", "����", "ɽ��", _
    "����", "����", "����", "�㶫", "����", "����", "����", "�Ĵ�", "����", "����", "����", "����", "����", "�ຣ", "����", "�½�")
    D = Array("ȫ����", "������", "�졡��", "�ӡ���", "ɽ����", "�ɡ���", "������", "�ϡ���", "������", "�㡡��", "������", "������", "������", "ɽ����", _
    "�ӡ���", "������", "������", "�㡡��", "�㡡��", "������", "�ء���", "�ġ���", "����", "�ơ���", "������", "�¡���", "�ʡ���", "�ࡡ��", "������", "�¡���")
    
    For i = 0 To 29
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = D(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i

End Sub

Sub A01_ʡ���滻()

    Dim C As Variant
    Dim D As Variant
    C = Array("�� ��", "�� ��", "�� ��", "�� ��", "ɽ ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "ɽ ��", _
    "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��", "�� ��")
    D = Array("�ܼ�", "����", "���", "�ӱ�", "ɽ��", "����", "����", "�Ϻ�", "����", "�㽭", "����", "����", "����", "ɽ��", _
    "����", "����", "����", "�㶫", "����", "����", "����", "�Ĵ�", "����", "����", "����", "����", "����", "�ຣ", "����", "�½�")
    
    For i = 0 To 29
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = D(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i

End Sub
Sub ����Ҫ��()

A00_��ҳ��ʽ

Application.ScreenUpdating = False '�ر���Ļ����


'��ѡ����ı�������в���,���û��ѡ��,���Զ�ѡ�������ĵ�
Dim MyRange As Range
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If

    Set MyRange = Selection.Range
    
    '����һ��32��Ԫ�ص�����,Ԫ��Ϊ��ʡ����
    Dim C As Variant
    C = Array("����", "���", "�ӱ�", "ɽ��", "���ɹ�", "����", "����", "������", "�Ϻ�", "����", "�㽭", "����", "����", "����", "ɽ��", _
    "����", "����", "����", "�㶫", "����", "����", "����", "�Ĵ�", "����", "����", "����", "����", "����", "�ຣ", "����", "�½�", "�½������������")
   
   '��һ���Ҳ����и�ʽת��,��ʡ�ݽ�������Ӵ�,����Ӧ���ӿ���
   For i = 0 To 31
    With Selection.Find
        .ClearFormatting
        .Execute FindText:=C(i) & "^p"
        If .Found = True Then
            .Parent.Expand Unit:=wdParagraph
            
    '����Ӵ�
    Selection.Paragraphs(1).Range.Font.Bold = True
    OLD = Selection.Paragraphs(1).Range.Text
    
    'ɾ����ǰ�ո�,��ǰ�Ͷκ�����һ����
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
 Application.ScreenUpdating = True '�ָ���Ļ����

End Sub
Sub ҳ������()

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
Sub ȱʡҳ������()

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

Sub ҳ���()
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
Sub �ı�X()
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
        .Replacement.Text = "^p����"
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
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
     
Else
    MsgBox "��ע�⡿����㲻�ڱ���У�" & Chr(13) & _
           "�뽫�����ŵ��������ⵥԪ���У� " & Chr(13) & _
           "Ȼ����ִ�б��꣬лл��"
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

Sub ����()
Selection.HomeKey Unit:=wdStory

'�������ŵĵ�һ���͵�2�����ֱ���Ϊy1,y2
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

'�滻Ϊ��ȷ����������
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
Sub ɾ������()


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

Sub ճ���ı�()
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
' ȡ�õ�ǰ�ļ���
    Set MyDOC = Application.ActiveWindow.Document
'    MsgBox myDoc
    NM = Left(MyDOC, Len(MyDOC) - 4)
    ActiveDocument.SaveAs FileName:=NM & "1" & ".htm", FileFormat:=wdFormatHTML

  '      Application.DisplayAlerts = wdAlertsNone
  '  ActiveDocument.Close SaveChanges:=wdSaveChanges
     ActiveWindow.Close wdDoNotSaveChanges

        Next i
    Else
        MsgBox "û�ҵ������ĵ�"
    End If
End With
 
End Sub

Sub ���Ŀո�()
    Selection.TypeText Text:="��"
End Sub
Sub ���ת��()

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
Sub �����ĸ�1()
    Dim NM As String
    Dim P As Integer
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    Application.Run MacroName:="Normal.NewMacros.ҳ������"

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
    'Application.Run MacroName:="Normal.NewMacros.�س��滻"

    ActiveDocument.Tables(1).Columns(1).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(1).PreferredWidth = 40
    ActiveDocument.Tables(1).Columns(2).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(2).PreferredWidth = 60
    Selection.HomeKey Unit:=wdStory
    'Application.Run MacroName:="Normal.NewMacros.ɾ���ո�K"

End Sub
Sub �ı����Ʊ�λ()

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

Sub ת��Ӣ�ı��()
    Dim MyRange As Range
    Set MyRange = Selection.Range
    
    With MyRange
        pn = MyRange.Paragraphs.Count
        For i = 1 To pn
            T = Trim(MyRange.Paragraphs(i).Range.Text)
            L = Len(T)

        '��"|"�滻Ϊ�Ʊ�λ Chr(9)
        C = "|"
        T = Trim(MyRange.Paragraphs(i).Range.Text)
        L = Len(T)
        P = InStr(1, T, C, 1)
        S1 = Trim(Left(T, P - 1))
        S2 = Trim(Right(T, L - P))

        '�жϴ�д��ĸ����ɾ����д��ĸǰ������
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
Sub ���ֵ�һ��1()
    Dim MyRange As Range
    Set MyRange = Selection.Range
    Set MyRange1 = ActiveDocument _
    .Range(start:=ActiveDocument.Content.End - 1)

    
    With MyRange
        pn = MyRange.Paragraphs.Count
        For i = 1 To pn

        '��"|"�滻Ϊ�Ʊ�λ Chr(9)
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

Sub A01_ɾ����()

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
Sub ת�����ֲ���1()
    
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

        '��"|"�滻Ϊ�Ʊ�λ Chr(9)
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
        .NameFarEast = "����"
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

Sub ���б��Ķ��ߺ͵��߼Ӵ�()

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
Sub �����()
    
    T = Selection.Range.Text
    If Len(T) = 0 Then
    MsgBox "����ѡ������ı�"
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
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .name = "Times New Roman"
        .Size = 12
        .Bold = True
    End With
    End If
End Sub

Sub ���ֵ�һ��()
    Dim MyRange As Range
    Set MyRange = Selection.Range
    Set MyRange1 = ActiveDocument.Range(start:=ActiveDocument.Content.End - 1)
    
    With MyRange
        pn = MyRange.Paragraphs.Count
        For i = 1 To pn

        '��"|"�滻Ϊ�Ʊ�λ Chr(9)
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

Sub ת�����������ֲ���()
    
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
        '��"|"�滻Ϊ�Ʊ�λ Chr(9)
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
        Application.Run MacroName:="Normal.NewMacros.ճ���ı�"
        PN1 = ActiveDocument.Content.Paragraphs.Count
            Set MyRange2 = ActiveDocument _
    .Range(start:=0, End:=ActiveDocument.Content.Paragraphs(PN1 - 1).Range.End)
    MyRange2.Select
        Application.Run MacroName:="Normal.NewMacros.ת�����ֲ���"
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
        Application.Run MacroName:="Normal.NewMacros.�س��滻"
        ActiveDocument.Content.Tables(1).Range.Select
        Selection.Copy
        ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
        MyDOC.Activate
        Selection.Delete
        Selection.Paste
        ActiveDocument.Content.Tables(1).Columns(1).Cells(1).Select
        Selection.InsertRowsAbove 2

End Sub
Sub ת�����ֲ���()
    
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

Sub ����ĸ��д()
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
    
    '���Сд
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
Sub ���Сд()
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

Sub ȫ��Сд()
    Dim RNG As Range
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If
    Set RNG = Selection.Range
    With RNG
        .Text = LCase(RNG.Text)
    End With
        
End Sub
Sub ȫ����д()
    Dim RNG As Range
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If
    Set RNG = Selection.Range
    With RNG
        .Text = UCase(RNG.Text)
    End With
        
End Sub

Sub �ı�W()
    '����������;��ת���ı����趨��ʽ
    
    '���ĵ�����ת����Ϊ���ı�
    Selection.WholeStory
    Selection.Cut
    Selection.Collapse Direction:=wdCollapseStart
    Selection.Range.PasteSpecial DataType:=wdPasteText
    
    '���������Ϊ������壬�м���趨Ϊ�̶�16��
    Selection.HomeKey Unit:=wdStory
    Selection.TypeParagraph
    Selection.WholeStory
        Selection.EndKey Unit:=wdStory
    Selection.WholeStory
    With Selection.Font
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .name = "Times New Roman"
        .Size = 10.5
    End With
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 16
    End With

    'ת������س���ΪӲ�س���ɾ���ո񣬶������
    Dim O As Variant
    Dim R As Variant
    O = Array("^l", " ", "^p^p", "��", "^p", "����^p")
    R = Array("^p", "", "^p", "", "^p����", "")
    
    For i = 0 To 5
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = O(i)
        .Replacement.Text = R(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    
    Selection.WholeStory

'�����мӴ�
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
    
    '��ȫ������ת��Ϊ�������
    Selection.WholeStory
    Dim C As Variant
    Dim D As Variant
    C = Array("��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��")
    D = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".")
    For i = 0 To 10
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = C(i)
        .Replacement.Text = D(i)
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    

    'ÿ������֮���һ��
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

Sub ����()
    '���û��ѡ���ı�����ʾ�û�ѡ��
    If Len(Selection.Range.Text) = 0 Then
        ActiveDocument.Paragraphs(1).Range.Select
'        MsgBox "��ע�⡿û��ѡ����Ϊ����ĵ��ļ������ı�" & Chr(13) & _
'           "��ѡ�������ļ�������ѡ�У� " & Chr(13) & _
'           "Ȼ����ִ�б��꣬лл��"
        Else
            Selection.Copy
            Set MyDOC = Application.ActiveWindow.Document
            ' �½�һ���հ��ĵ�
            Documents.Add DocumentType:=wdNewBlankDocument
            Set MyDocN = Application.ActiveWindow.Document
            Application.Run MacroName:="ճ���ı�"
            Application.Run MacroName:="�س��滻"
            Application.Run MacroName:="ɾ���ո�K"
            Application.Run MacroName:="ɾ�����Ŀո�"
            NM = Left(MyDocN.Paragraphs(1).Range.Text, Len(MyDocN.Paragraphs(1).Range.Text) - 1)
            MyDOC.Activate
            ChangeFileOpenDirectory "D:\00 F2013\02 ��������\"
            ActiveDocument.SaveAs FileName:=NM & ".doc", FileFormat:=wdFormatDocument
            MyDocN.Close SaveChanges:=wdDoNotSaveChanges
            
    End If
        ChangeFileOpenDirectory "D:\00 F2013\"
        Documents(NM & ".doc").Activate


End Sub
Sub A01_�����()

    ChangeFileOpenDirectory "D:\00 F2013\02 ��������"

    NM1 = ActiveDocument.Paragraphs(1).Range.Text
    NM1 = Left(NM1, Len(NM1) - 1)
    
    NM2 = ActiveDocument.Shapes(1).TextFrame.TextRange.Text
    NM2 = Left(NM2, Len(NM2) - 1)
    
    NM = NM1 & NM2
    ActiveDocument.SaveAs FileName:=NM & ".doc", FileFormat:=wdFormatDocument
    ChangeFileOpenDirectory "D:\00 F2013\"
    
End Sub

Sub A01_�����1()

    ChangeFileOpenDirectory "D:\00 F2013\02 ��������"

    NM1 = ActiveDocument.Paragraphs(3).Range.Text
    NM1 = Left(NM1, Len(NM1) - 1)
    
    NM2 = ActiveDocument.Shapes(1).TextFrame.TextRange.Text
    NM2 = Left(NM2, Len(NM2) - 1)
    
    NM = NM1 & NM2
    ActiveDocument.SaveAs FileName:=NM & ".doc", FileFormat:=wdFormatDocument
    ChangeFileOpenDirectory "D:\00 F2013\"
    
End Sub


Sub ���뷭����Ƶ�()

    ZS = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharacters)
    ZS1 = Round(ZS / 1000, 2)
    If ZS1 < 1 Then
        qzs = "0" & ZS1
        Else
        qzs = ZS1
    End If
    
        ChangeFileOpenDirectory "D:\F2006\������"
    
    'ȷ���ļ���
    
    Application.DisplayAlerts = wdAlertsNone
    M = IIf(Month(Date) < 10, "0" & Month(Date), Month(Date))
    D = IIf(Day(Date) < 10, "0" & Day(Date), Day(Date))
    y = Year(Date)
    N = "G" & y & M & D & "-"

Set fs = Application.FileSearch
With fs
    .LookIn = "D:\F2006\������"
    .FileName = N
    
    If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
        FN = N & 1
    Else
        FN = N & .FoundFiles.Count + 1
    End If
End With
'MsgBox Fn

    Selection.HomeKey Unit:=wdStory
    Selection.InsertFile FileName:="C:\00 Word_dot\���빤�����̿��Ƶ�.doc", ConfirmConversions:=False
    ActiveDocument.Tables(1).Cell(1, 2).Select
    Selection.TypeText Text:=FN
    ActiveDocument.Tables(1).Cell(1, 4).Select
    T = ActiveDocument.Tables(2).Cell(1, 1).Range.Text
    nt = Left(T, Len(T) - 2)
    Selection.TypeText Text:=nt
    ActiveDocument.Tables(1).Cell(1, 6).Select
    
    Selection.TypeText Text:=qzs
    ActiveDocument.Tables(1).Cell(5, 2).Select
    Selection.TypeText Text:=Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
    ActiveDocument.Tables(1).Cell(6, 2).Select
    Selection.TypeText Text:=Year(Date) & "��" & Month(Date) & "��" & Day(Date) + 1 & "��"
    
    Application.Run MacroName:="Normal.NewMacros.ҳ������"

    ActiveDocument.SaveAs FileName:=FN, FileFormat:=wdFormatDocument

End Sub

Sub ��ѡ�����ʽ��׼��()
    
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
        S1 = "����" & T
        MyPara.Range.Text = S1
     Next MyPara
     End With

End Sub
Sub �ı���ʽ()
    Application.Run MacroName:="ɾ��ǰ��"
    Application.Run MacroName:="ɾ����"
    Application.Run MacroName:="��ǰ�ӿ�"
    Selection.WholeStory
    'Selection.Style = ActiveDocument.Styles("C����")
    Application.Run MacroName:="�ӿ���"
        Selection.WholeStory
        Selection.HomeKey Unit:=wdStory
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Application.Run MacroName:="�������C"
        Selection.WholeStory
        Selection.EndKey Unit:=wdStory
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.Delete Unit:=wdCharacter, Count:=1
End Sub
Sub ����Զ�����()
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
'    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.MoveRight Unit:=wdCharacter, Count:=1

'    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
        Selection.MoveRight Unit:=wdCharacter, Count:=1

'    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
End Sub
Sub ���ļӴ�()

'�����ض����ַ������磺"��"��Ȼ�󽫸��ַ���ǰ�����ݼӴ�
    T = "����"
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
    
        Z = "�¡�"
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


End Sub
Sub ���()

    '�԰�ÿһ�ж�����һ��������ĵ���׼����ɾ������Ķ�����
    Selection.WholeStory
    
    '��һ�������Ǻ���������Ŀո�ĵط�����Ϊ"��"
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p����"
        .Replacement.Text = "��"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    'ɾ�����ж�����
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = ""
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
   '��"��"��ת��Ϊ������
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "��"
        .Replacement.Text = "^p����"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub
Sub ����˳��ߵ�()

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
    
    Application.Run MacroName:="ɾ����"

End Sub

Sub MyFilePrint()
pass$ = InputBox("�������ӡ���룺")
If pass$ = "abcd" Then
Dialogs(wdDialogFilePrint).Show
DName = ActiveDocument.Path + "\" + ActiveDocument.name
If ActiveDocument.Path = "" Then DName = "δ�����ĵ�"
Tim = str(Date) + " �� " + str(Time)
Open "c:\print.txt" For Append As #1
Print #1, "�� " + Tim + " ��ӡ " + DName
Close #1
Else
MsgBox ("����������������Ա��ϵ��")
End If
End Sub


Sub CPFL()
Dim SourceFile, DestinationFile
SourceFile = "C:\00 Word_Dot\richeng.doc"    ' ָ��Դ�ļ�����
DestinationFile = "C:\00 Word_Dot\richeng.bak"    ' ָ��Ŀ���ļ�����
FileCopy SourceFile, DestinationFile    ' ��Դ�ļ������ݸ��Ƶ�Ŀ���ļ��С�
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
                
Sub �����ĸ�()

    Dim NM As String
    Dim P, pn As Integer
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    Application.Run MacroName:="Normal.NewMacros.ҳ������"

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
    'Application.Run MacroName:="Normal.NewMacros.�س��滻"

    ActiveDocument.Tables(1).Columns(1).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(1).PreferredWidth = 40
    ActiveDocument.Tables(1).Columns(2).PreferredWidthType = wdPreferredWidthPercent
    ActiveDocument.Tables(1).Columns(2).PreferredWidth = 60
    Selection.HomeKey Unit:=wdStory
    'Application.Run MacroName:="Normal.NewMacros.ɾ���ո�K"

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
'    Application.Run macroName:="Normal.NewMacros.���뷭����Ƶ�"
    

End Sub

Sub macro00()
    A00_��ҳ��ʽ
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
    
    Application.Run MacroName:="Normal.NewMacros.���뷭����Ƶ�"

 End Sub
 
 Sub Macro0()
 
    A00_��ҳ��ʽ
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
    Application.Run MacroName:="Normal.NewMacros.ҳ�߾�2����"
    
'    Application.Run MacroName:="Normal.NewMacros.���뷭����Ƶ�"
 
 End Sub
 
Sub FNtest()
    Application.DisplayAlerts = wdAlertsNone
    N = "F" & Month(Date) & Day(Date) & "-"
    j = 1

Set fs = Application.FileSearch
With fs
    .LookIn = "D:\00 F2005\04 ������"
    .FileName = N & j
    
    If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
        FN = N & j
    Else
        FN = N & .FoundFiles.Count + 1
    End If
End With
'MsgBox Fn
 
End Sub
Sub Ŀ¼()
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
' ȡ�õ�ǰ�ļ���
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
        MsgBox "û�ҵ������ĵ�"
    End If
End With
 
End Sub


Sub ������()

    ChangeFileOpenDirectory "C:\Users\thtfpc\Desktop"
    
    'ȷ���ļ���
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

'ChangeFileOpenDirectory "C:\Documents and Settings\ZLG005\����"

Selection.WholeStory
Selection.HomeKey Unit:=wdStory
ActiveDocument.Tables(1).Columns(2).Select
Selection.Cut

    ActiveDocument.SaveAs FileName:=FNM & ".htm", FileFormat:=wdFormatHTML

    Documents.Close SaveChanges:=wdDoNotSaveChanges
    Documents.Add DocumentType:=wdNewBlankDocument
ChangeFileOpenDirectory "D:\00 F2006\"

End Sub

Sub ת�����()
    
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
Sub ����ת��Ϊ�Ʊ�λ()

'*****************************************************
'��������Ĺ��ܣ��������滻Ϊ�Ʊ�λ����ɾ������Ŀո�
'���ߣ������
'���ڣ�2005��9��7��
'*****************************************************

    '����һ����Χ��Range������
    Dim MyRange As Range
    '���û��ѡ��Χ����ָ����ΧΪ�����ĵ�
    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    '�趨��Χ����Ϊѡ��ķ�Χ
    Set MyRange = Selection.Range
    'ȡ����ѡ����Ķ�������
    pn = MyRange.Paragraphs.Count

        '�趨�ָ��ַ�
        txt1 = "|"
        '��ѡ��ķ�Χ������ز���
        With MyRange
            '�趨һ������ѭ����jΪ����������ѡ����Χ�ĵ�һ�����俪ʼ��ѭ�������һ������
            For j = 1 To pn
                '�������t����������е��ı�(�������س���
                t1 = MyRange.Paragraphs(j).Range.Text
                T = Left(t1, Len(t1) - 1)
                '�������p2Ϊ"|"��һ���ڶ����г��ֵ�λ�ã����������û�У���p2=0
                p2 = InStr(1, T, txt1, 1)
                '�������s2�������ı����ܳ��ȣ�������ǰ���ո�
                S2 = LTrim(T)
                
                '����һ��ѭ��������ı�����"|"�����ҵ���λ�ã����������Ʊ�λChr(9)�滻��ֱ��ȫ���滻Ϊֹ
                Do Until p2 = 0
                
                    '�ж��ı����Ƿ���"|"
                    If InStr(1, S2, txt1, 1) > 0 Then
                        p2 = InStr(1, S2, txt1, 1)
                        '����S3Ϊ����ǰ���ı���S4Ϊ���ߺ���ı�
                        S3 = Trim(Left(S2, p2 - 1))
                        S4 = Right(S2, Len(S2) - p2)
                    End If
                    S2 = Trim(S3) & Chr(9) & Trim(S4)
                    p2 = InStr(1, S2, txt1, 1)
                Loop
            '�滻�����е��ı�
            MyRange.Paragraphs(j).Range.Text = S2 & Chr(13)
        '��ת����һ������
        Next j
    End With
End Sub

Sub ת��Ϊ��è�ļ�()

If Selection.Information(wdWithInTable) = True Then
    Selection.Copy
  Else
    MsgBox "��ע�⡿����㲻�ڱ���У�" & Chr(13) & _
           "���������뽫�����ŵ��������ⵥԪ" & Chr(13) & _
           "�����������У� Ȼ����ִ�б��꣬лл��"
   End If

' �½�һ���հ��ĵ�
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
    ct = "��"
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
Sub JS��ӡ�к���()

If Selection.Information(wdWithInTable) = True Then
    Selection.Copy
  Else
    MsgBox "��ע�⡿����㲻�ڱ���У�" & Chr(13) & _
           "����������ѡ���������һ��" & Chr(13) & _
           "��������Ȼ����ִ�б��꣬лл��"
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
    
    ' �½�һ���հ��ĵ�
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

Sub ɾ�����е�����()
    
'------------------------------------------------------------------
'��������Ĺ��ܣ�ɾ�������ڵ����ݣ���ɾ������Ŀո񼰶�ǰ�κ�ո�
'���ߣ������
'���ڣ�2005��11��25��
'------------------------------------------------------------------

    '����һ����Χ��Range������
    Dim MyRange As Range
    '���û��ѡ��Χ����ָ����ΧΪ�����ĵ�
    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    '�趨��Χ����Ϊѡ��ķ�Χ
    Set MyRange = Selection.Range
    'ȡ����ѡ����Ķ�������
    pn = MyRange.Paragraphs.Count

        '�趨�ָ��ַ�
'        txt1 = Chr(40) '(
'        txt2 = Chr(41) ')
        txt1 = Chr(60) '<
        txt2 = Chr(62) '>
'        txt1 = Chr(123) '{
'        txt2 = Chr(125) '}
        
        '��ѡ��ķ�Χ������ز���
        With MyRange
            '�趨һ������ѭ����jΪ����������ѡ����Χ�ĵ�һ�����俪ʼ��ѭ�������һ������
            For j = 1 To pn
                '�������t����������е��ı�(�������س���
                t1 = MyRange.Paragraphs(j).Range.Text
                T = Left(t1, Len(t1) - 1)
                '��������������p1,p2,�ֱ����"("��")"��һ���ڶ����г��ֵ�λ�ã����������û�У���Ϊ0
                P1 = InStr(1, T, txt1, 1)
                p2 = InStr(1, T, txt2, 1)
                '�������s2�������ı����ܳ��ȣ�������ǰ���ո�
                S2 = LTrim(T)
                
                '����һ��ѭ��������ı�����("��")"�����ҵ���λ�ã���ɾ֮��ֱ��ȫ��ɾ��Ϊֹ
                Do Until P1 = 0 Or p2 = 0
                
                    '�ж��ı����Ƿ���("��")"
                    If InStr(1, S2, txt1, 1) > 0 And InStr(1, S2, txt2, 1) > 0 Then
                        P1 = InStr(1, S2, txt1, 1)
                        p2 = InStr(1, S2, txt2, 1)
                        If p2 < P1 Then
                        p2 = InStr(P1, S2, txt2, 1)
                       End If
                        '����S3Ϊ("ǰ���ı���S4Ϊ")"����ı�
                        S3 = Trim(Left(S2, P1 - 1))
                        S4 = Right(S2, Len(S2) - p2)
                    End If
                    S2 = Trim(S3) & Trim(S4)
                    P1 = InStr(1, S2, txt1, 1)
                    p2 = InStr(1, S2, txt2, 1)
                Loop
            '�滻�����е��ı�
            MyRange.Paragraphs(j).Range.Text = S2 & Chr(13)
        '��ת����һ������
        Next j
    End With
    Application.Run MacroName:="Normal.NewMacros.EDC"
    Application.Run MacroName:="Normal.NewMacros.ɾ����"

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

Sub ��ASCII��()
    If Len(Selection.Range.Text) = 0 Then
        MsgBox "��ע�⡿����ѡ���ַ���Ȼ����ִ�б������лл��"
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
Sub ����ɾ��һϵ���ַ�()
    Selection.WholeStory
    Dim A, B As Variant
    A = Array(Chr(21), Chr(22), Chr(23))
    For i = 0 To UBound(A)
        ɾ��ĳ���ַ� (A(i))
    Selection.WholeStory
Next i
End Sub
Function ɾ��ĳ���ַ�(txt1)
    '����һ����Χ��Range������
    Dim MyRange As Range
    Selection.WholeStory
    If Len(Selection.Range.Text) = 0 Then
        MsgBox "��ע�⡿����ѡ��Ȼ����ִ�б������лл��"
    Else
    Set MyRange = Selection.Range
'        txt1 = Chr(41)  '����Ҫɾ�����ַ�
        T = MyRange.Text
        p2 = InStr(1, T, txt1, 1)
        S2 = T
       Do Until p2 = 0
        If InStr(1, S2, txt1, 1) > 0 Then
           p2 = InStr(1, S2, txt1, 1)
           S3 = Trim(Left(S2, p2 - 1))  '����S3Ϊ����ǰ���ı�
           S4 = Right(S2, Len(S2) - p2) 'S4Ϊ���ź���ı�
         End If
           S2 = S3 & S4
           p2 = InStr(1, S2, txt1, 1)
        Loop
            MyRange.Text = S2
    End If

End Function
Function �滻ĳ���ַ�(TC)
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
           S3 = Trim(Left(S2, p2 - 1))  '����S3Ϊ����ǰ���ı�
           S4 = Right(S2, Len(S2) - p2 - 1) 'S4Ϊ���ź���ı�
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
        Application.Run MacroName:="Normal.NewMacros.ɾ�����е�����"
        ActiveDocument.SaveAs FileName:=.FoundFiles(i), FileFormat:=wdFormatText, AddToRecentFiles:=False
        Document.Add
        Next i
    Else
        MsgBox "û�ҵ������ĵ�"
    End If
End With
 
End Sub
Sub ����ɾ�������е�����()
    D = "D:\txttest\"
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(D)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.name
        s = Left(s, Len(s) - 4)
        ChangeFileOpenDirectory D
    '���ļ�
        Documents.Open FileName:=s & ".txt", ConfirmConversions:=False, _
        ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        WritePasswordTemplate:="", Format:=wdOpenFormatAuto, Encoding:=936
    '���к�����
'       Application.Run MacroName:="Normal.NewMacros.A����"
        Application.Run MacroName:="Normal.NewMacros.ɾ�����е�����"
        Application.Run MacroName:="Normal.NewMacros.ɾ��ǰ��"
        Application.Run MacroName:="Normal.NewMacros.����"
        Application.Run MacroName:="Normal.NewMacros.ɾ����"
    '�����ļ����ļ�������ˡ�n��
        ActiveDocument.SaveAs FileName:="N_" & s, FileFormat:=wdFormatText, _
        LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
        ActiveWindow.Close
    Next
End Sub
Function dc(C) 'ɾ���ĵ�������ָ�����ַ��� ����
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
    dc ("��")
   ' DC (Chr(13))
End Sub



Sub jf()
'*****************************************************
'��������Ĺ��ܣ����ĵ���ѡ������ת��Ϊ����
'���ߣ������
'���ڣ�2005��11��25��
'*****************************************************

ML = "D:\00 Word_Dot\"
ChangeFileOpenDirectory ML


Documents.Open ("jtchar.doc")
jt = ActiveDocument.Content.Paragraphs(1).Range.Text
ActiveDocument.Close

Documents.Open ("ftchar.doc")
ft = ActiveDocument.Content.Paragraphs(1).Range.Text
ActiveDocument.Close

    '����һ����Χ��Range������
    Dim MyRange As Range
    '���û��ѡ��Χ����ָ����ΧΪ�����ĵ�
    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    '�趨��Χ����Ϊѡ��ķ�Χ
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
'��������Ĺ��ܣ����ĵ���ѡ������ת��Ϊ����
'���ߣ������
'���ڣ�2005��11��25��
'*****************************************************

ML = "D:\00 Word_Dot\"
ChangeFileOpenDirectory ML


Documents.Open ("jtchar.doc")
jt = ActiveDocument.Content.Paragraphs(1).Range.Text
ActiveDocument.Close

Documents.Open ("ftchar.doc")
ft = ActiveDocument.Content.Paragraphs(1).Range.Text
ActiveDocument.Close

    '����һ����Χ��Range������
    Dim MyRange As Range
    '���û��ѡ��Χ����ָ����ΧΪ�����ĵ�
    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    '�趨��Χ����Ϊѡ��ķ�Χ
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

'SubStr()    ���Ļ�ȡ���ִ������Mid()
'Strlen()    ���Ļ��ִ����ȣ����Len()
'StrLeft()   ���Ļ�ȡ���ִ������Left()
'StrRight()  ���Ļ�ȡ���ִ������Right()
'isChinese() Checkĳ�����Ƿ�������

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

Function shc(txt1)    '����һ����Χ��Range������
    Dim MyRange As Range
    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    
    Set MyRange = Selection.Range
'        txt1 = Chr(41)  '����Ҫɾ�����ַ�
        T = MyRange.Text
        p2 = InStr(1, T, txt1, 1)
        S2 = T
       Do Until p2 = 0
        If InStr(1, S2, txt1, 1) > 0 Then
           p2 = InStr(1, S2, txt1, 1)
           S3 = Trim(Left(S2, p2 - 1))  '����S3Ϊ����ǰ���ı�
           S4 = Right(S2, Len(S2) - p2) 'S4Ϊ���ź���ı�
         End If
           S2 = S3 & S4
           p2 = InStr(1, S2, txt1, 1)
        Loop
            MyRange.Text = S2

End Function
Sub ת��NJ98()
    '����һ����Χ��Range������
    Dim MyRange As Range
    If Len(Selection.Range.Text) = 0 Then
    Selection.WholeStory
    End If
    Set MyRange = Selection.Range
        txt1 = Chr(124)  '����Ҫɾ�����ַ�
        txt2 = Chr(9)
        T = MyRange.Text
        p2 = InStr(1, T, txt1, 1)
        S2 = T
       Do Until p2 = 0
        If InStr(1, S2, txt1, 1) > 0 Then
           p2 = InStr(1, S2, txt1, 1)
           S3 = Trim(Left(S2, p2 - 1))  '����S3Ϊ����ǰ���ı�
           S4 = Trim(Right(S2, Len(S2) - p2)) 'S4Ϊ���ź���ı�
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
        Application.Run MacroName:="Normal.NewMacros.����"
        Application.Run MacroName:="Normal.NewMacros.ɾ����"
        
        Selection.WholeStory
        A2 = "�� �� Beijing"
        A3 = "�� �� Tianjin"
        TXT = Selection.Range.Text
        P1 = InStr(1, TXT, A2, vbTextCompare)
        p2 = InStr(1, TXT, A3, vbTextCompare)
        If P1 > 0 And p2 > 0 Then
        Application.Run MacroName:="Normal.NewMacros.ʡ���滻"
        End If
        
        shc ("ed by Region")


End Sub

Sub ʡ���滻()
Dim A, B As Variant
A1 = "ȫ �� National Total"
A2 = "�� �� Beijing"
A3 = "�� �� Tianjin"
A4 = "�� �� Hebei"
A5 = "ɽ �� Shanxi"
A6 = "���ɹ� Inner Mongolia"
A7 = "�� �� Liaoning"
A8 = "�� �� Jilin"
A9 = "������ Heilongjiang"
A10 = "�� �� Shanghai"
A11 = "�� �� Jiangsu"
A12 = "�� �� Zhejiang"
A13 = "�� �� Anhui"
A14 = "�� �� Fujian"
A15 = "�� �� Jiangxi"
A16 = "ɽ �� Shandong"
A17 = "�� �� Henan"
A18 = "�� �� Hubei"
A19 = "�� �� Hunan"
A20 = "�� �� Guangdong"
A21 = "�� �� Guangxi"
A22 = "�� �� Hainan"
A23 = "�� �� Chongqing"
A24 = "�� �� Sichuan"
A25 = "�� �� Guizhou"
A26 = "�� �� Yunnan"
A27 = "�� �� Tibet"
A28 = "�� �� Shaanxi"
A29 = "�� �� Gansu"
A30 = "�� �� Qinghai"
A31 = "�� �� Ningxia"
A32 = "�� �� Xinjiang"
A33 = "���ֵ��� Not Classifi-"

B1 = "ȫ���ϼ�" & Chr(9) & "National Total"
B2 = "������" & Chr(9) & "Beijing"
B3 = "�졡��" & Chr(9) & "Tianjin"
B4 = "�ӡ���" & Chr(9) & "Hebei"
B5 = "ɽ����" & Chr(9) & "Shanxi"
B6 = "���ɹ�" & Chr(9) & "Inner Mongolia"
B7 = "�ɡ���" & Chr(9) & "Liaoning"
B8 = "������" & Chr(9) & "Jilin"
B9 = "������" & Chr(9) & "Heilongjiang"
B10 = "�ϡ���" & Chr(9) & "Shanghai"
B11 = "������" & Chr(9) & "Jiangsu"
B12 = "�㡡��" & Chr(9) & "Zhejiang"
B13 = "������" & Chr(9) & "Anhui"
B14 = "������" & Chr(9) & "Fujian"
B15 = "������" & Chr(9) & "Jiangxi"
B16 = "ɽ����" & Chr(9) & "Shandong"
B17 = "�ӡ���" & Chr(9) & "Henan"
B18 = "������" & Chr(9) & "Hubei"
B19 = "������" & Chr(9) & "Hunan"
B20 = "�㡡��" & Chr(9) & "Guangdong"
B21 = "�㡡��" & Chr(9) & "Guangxi"
B22 = "������" & Chr(9) & "Hainan"
B23 = "�ء���" & Chr(9) & "Chongqing"
B24 = "�ġ���" & Chr(9) & "Sichuan"
B25 = "����" & Chr(9) & "Guizhou"
B26 = "�ơ���" & Chr(9) & "Yunnan"
B27 = "������" & Chr(9) & "Tibet"
B28 = "�¡���" & Chr(9) & "Shaanxi"
B29 = "�ʡ���" & Chr(9) & "Gansu"
B30 = "�ࡡ��" & Chr(9) & "Qinghai"
B31 = "������" & Chr(9) & "Ningxia"
B32 = "�¡���" & Chr(9) & "Xinjiang"
B33 = "���ֵ���" & Chr(9) & "Not Classified by Region"
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

Sub ����ת��NJ98()
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
    '���ļ�
        Documents.Open FileName:=s & ".txt", ConfirmConversions:=False, _
        ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        WritePasswordTemplate:="", Format:=wdOpenFormatAuto, Encoding:=936
    '���к�����
'       Application.Run MacroName:="Normal.NewMacros.A����"
        Application.Run MacroName:="Normal.NewMacros.ת��NJ98"
    '�����ļ����ļ�������ˡ�C��
        ActiveDocument.SaveAs FileName:=s & "C", FileFormat:=wdFormatDocument, _
        LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
    '�����ļ����ļ�������ˡ�E��
        ActiveDocument.SaveAs FileName:=s & "E", FileFormat:=wdFormatDocument, _
        LockComments:=False, Password:="", AddToRecentFiles:=False, WritePassword _
        :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
        
        Selection.WholeStory
        Application.Run MacroName:="Normal.NewMacros.ɾ����"
        ActiveDocument.Save
        ActiveWindow.Close
        
        Documents.Open (s & "C.doc")
        Selection.WholeStory
        Application.Run MacroName:="Normal.NewMacros.ɾ��Ӣ����ĸ"
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
        MsgBox "û�ҵ������ĵ�"
    End If
End With
 
End Sub
Sub ɾ��ո�()

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
Sub ���������()

    Set mytable = ActiveDocument.Tables(1)
    Set aColumn = mytable.Columns.Add(BeforeColumn:=mytable.Columns(1))
    For Each aCell In aColumn.Cells
        aCell.Range.Delete
        aCell.Range.InsertAfter Num + 1
        Num = Num + 1
    Next aCell
    
End Sub

Sub ɾ�����п���()
If ActiveDocument.Tables.Count >= 1 Then
    Set mytable = ActiveDocument.Tables(1)
    C = ActiveDocument.Tables(1).Columns.Count
    R = ActiveDocument.Tables(1).Rows.Count
'    MsgBox R & "��" & C
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
Sub �����ݵ���()
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
End Sub
Sub �̶��п�()
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitFixed)
End Sub
Sub �����ڵ���()
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
End Sub
Sub ҳ����ͼ()
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
End Sub
Sub ��ҳ��ͼ()
    ActiveWindow.View.Type = wdWebView
End Sub
Sub ��ͨ��ͼ()
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdNormalView
    Else
        ActiveWindow.View.Type = wdNormalView
    End If
End Sub
Sub ���()
        Application.Run MacroName:="Normal.NewMacros.ɾ�����Ŀո�"
        Application.Run MacroName:="Normal.NewMacros.ҳ����ͼ"
        Application.Run MacroName:="Normal.NewMacros.�����ݵ���"
        Application.Run MacroName:="Normal.NewMacros.�̶��п�"

End Sub
Sub ȫ�Ǳ��ת����Ǳ��()

    'ȫ�Ǳ�����ת����Ǳ�����
    Selection.WholeStory
    Dim C As Variant
    Dim D As Variant
    C = Array("��", "��", "��", "��", "��", "��", "��", "��", "��", Chr(37), Chr(-23643), "��", "��", "��")
    
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

Sub ���볣��ͳ������()
    
    Dim MyRange As Range
    Dim f, R As Variant
    
    Dim CH(), EN() As Variant
        If Len(Selection.Range.Text) = 0 Then
        Selection.WholeStory
        End If
    Set MyRange = Selection.Range
    
    Documents.Open FileName:="D:\00 Word_Dot\����ͳ������Ӣ�����ձ�.doc"
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
    Application.Run MacroName:="Normal.NewMacros.����"
    Application.Run MacroName:="Normal.NewMacros.ɾ�κ��"
    Application.Run MacroName:="Normal.NewMacros.ɾ�κ��"
    Application.Run MacroName:="Normal.NewMacros.ɾ����"
    Application.Run MacroName:="Normal.NewMacros.����_��Ԫ"
    Application.Run MacroName:="Normal.NewMacros.����_���"
    Application.Run MacroName:="Normal.NewMacros.����_��ƽ����"
    Application.Run MacroName:="Normal.NewMacros.ȫ�Ǳ��ת����Ǳ��"
    Application.Run MacroName:="Normal.NewMacros.����"
    Application.Run MacroName:="Normal.NewMacros.ɾ����"

    
End Sub

Sub ����_��Ԫ()

    N = ActiveDocument.Paragraphs.Count
    For i = 1 To N
        T = ActiveDocument.Paragraphs(i).Range.Text
        T = Left(T, Len(T) - 1)
        t1 = "��Ԫ"
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

Sub ����_���()

    N = ActiveDocument.Paragraphs.Count
    For i = 1 To N
        T = ActiveDocument.Paragraphs(i).Range.Text
        T = Left(T, Len(T) - 1)
        t1 = "���"
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

Sub ����_��ƽ����()

    N = ActiveDocument.Paragraphs.Count
    For i = 1 To N
        T = ActiveDocument.Paragraphs(i).Range.Text
        T = Left(T, Len(T) - 1)
        t1 = "��ƽ����"
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

Sub ������������()
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
        .NameFarEast = "����"
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
    MsgBox "��ע�⡿����㲻�ڱ���У�" & Chr(13) & _
           "���������뽫�����ŵ��������ⵥԪ" & Chr(13) & _
           "�����������У� Ȼ����ִ�б��꣬лл��"
   End If
End Sub
Sub �������뵥λ()

    '�ж��Ƿ�ѡ�����й�����
    If Len(Selection.Range.Text) = 0 Then
        MsgBox "��û��ѡ���κ����ݣ�" & Chr(13) & "��ѡ�С����ı��⡢��λ�ͱ��" & Chr(13) & "Ȼ�������б�������!"
    Else

    '�ж��Ƿ�ѡ�����й�����
    If Selection.Tables.Count = 0 Then
        MsgBox "��û��ѡ�б��" & Chr(13) & "��ѡ�С����ı��⡢��λ�ͱ��" & Chr(13) & "Ȼ�������б�������!"
    Else
    
    '����ѡ�е�������ȡ���ı���͵�λ���ı�
    Set MyRange = Selection.Range
    Title = MyRange.Paragraphs(1).Range.Text
    Title = Left(Title, Len(Title) - 1)
    Unit = MyRange.Paragraphs(2).Range.Text
    Unit = Left(Unit, Len(Unit) - 1)
    
    '���뵥λ�У������ø�ʽ
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
    
    '��������У������ø�ʽ
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
    
    'ɾ��ԭ�����ı�
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
Sub ��λ������()  '����������
    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(10), Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
    Selection.TypeText Text:=Chr(9) & "��������" & Chr(13)
    Selection.InsertDateTime DateTimeFormat:=Chr(9) & "EEEE��O��A��", InsertAsField:=False
End Sub
Sub ɾ����Ԫ���е�0()
'ָ������ı��Ϊ�ĵ��еĵ�1�ű�
Set mytable = ActiveDocument.Tables(1)
'�趨ѭ��
For Each celTable In mytable.Range.Cells
'�趨��Χ����������Ԫ�����ݵĻ��з�
    Set rngTable = ActiveDocument.Range(start:=celTable.Range.start, End:=celTable.Range.End - 1)
    If rngTable.Text = "0" Then
        rngTable.Text = ""
    End If
    Next celTable
End Sub

Sub GBE_table()
'
' GBE_table Macro
' ���� 2007-3-12 �� �����: ¼��
'
    Dim T As Integer
    T = ActiveDocument.Tables.Count ' ȡ���ĵ��б�������
    
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
        MsgBox "�ĵ���û�б���޷�ִ�б�������!"
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
'��������Ĺ��ܣ��Զ�����ͬ������������б�š����桢�Ǽ�
'���ߣ������
'���ڣ�2008��5��5��
'----------------------------------------------------------------------------
Sub G01_�Ǽ�()

    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Set MyDOC = Application.ActiveWindow.Document  'ȷ����ĵ�Ϊ�����ĵ�
    
    If MyDOC.Tables.Count = 0 Then  '�ж��ĵ����Ƿ��б��
        MsgBox ("�ĵ���û�б�񣬲���ִ�б������")
        mycheck = False
    Else
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        If TN <> "����ͳ�ƾ�������Ϣ���������" Then
            MsgBox ("��������Ϣ�������������ִ�б������")
            mycheck = False
        Else
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
            lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '��ȡ�������
            dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '��ȡ����������
            '�ж��Ƿ��Ѿ��Ǽ�
            If Len(bh) < 10 Then
                If IsDate(dt) Then
                    mycheck = True
                Else
                    MsgBox "�������ڸ�ʽ���������޸ģ���ִ�б����" & Chr(13) & "��ȷ�ĸ�ʽΪ��2008-5-1��2008-05-01��2008��5��1��"
                    mycheck = False
                End If
            Else
                MsgBox ("��������ѵǼǣ������ظ��Ǽǣ�")
                mycheck = False
            End If
        End If
    End If

    Do While mycheck = True
    
    Application.DisplayAlerts = wdAlertsNone
    M = IIf(Month(Date) < 10, "0" & Month(Date), Month(Date))
    D = IIf(Day(Date) < 10, "0" & Day(Date), Day(Date))
    y = Year(Date)
    If lb = "���˱�" Then
    N = "G" & y & M & D & "-"
    Else
    N = "D" & y & M & D & "-"
    End If

'�ڹ���Ŀ¼�в����Ƿ��е��������ļ������û�У���Ŵ�1��ʼ��������ԭ���ı�Ż������������
Set fs = Application.FileSearch
With fs
    .LookIn = "D:\��Ϣ����"
    .FileName = "����" & N
    
    If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
        FN = N & 1
    Else
        FN = N & .FoundFiles.Count + 1
    End If
End With

    If Len(ActiveDocument.Tables(1).Cell(2, 2).Range.Text) < Len(FN) Then
    ActiveDocument.Tables(1).Cell(2, 2).Select
    Selection.TypeText Text:=FN
    ActiveDocument.SaveAs FileName:="����" & FN, FileFormat:=wdFormatDocument   '�����ļ�
    End If
    
    If lb = "���˱�" Then
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

    Documents.Open FileName:="D:\��Ϣ����\�ǼǱ�1.doc"
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
'    MsgBox ("�ѳɹ��Ǽǣ�")
    End If
    
    If lb = "��λ��" Then
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

    Documents.Open FileName:="D:\��Ϣ����\�ǼǱ�2.doc"
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
'    MsgBox ("�ѳɹ��Ǽǣ�")
    End If
    
    Exit Do
    Loop
    
End Sub

'--------------------------------------------------------------
'��������Ĺ��ܣ���ȡ�������Ϣ���Զ����ɻ�ִ�����浽ָ��Ŀ¼
'���ߣ������
'���ڣ�2008��5��6��
'--------------------------------------------------------------

Sub G02_��ִ()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Set MyDOC = Application.ActiveWindow.Document  'ȷ����ĵ�Ϊ�����ĵ�
    If MyDOC.Tables.Count > 0 Then  '�ж��ĵ����Ƿ��б��
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        If TN = "����ͳ�ƾ�������Ϣ���������" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
            '�ж��Ƿ��Ѿ��Ǽ�
            If Len(bh) < 10 Then
                MsgBox ("�������δ�Ǽǣ����ȵǼǣ�")
                mycheck = False
            Else
                '�ж��Ƿ��Ѿ����ɹ��Ǽǻ�ִ
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\��Ϣ����"
                        .FileName = "��ִ" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("�Ѿ����ɹ��Ǽǻ�ִ�������ٴ����ɣ�")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("��������Ϣ�������������ִ�б������")
            mycheck = False
        End If
    Else
        MsgBox ("�ĵ���û�б�񣬲���ִ�б������")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '��ȡ���������
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '��ȡ��������λ����
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '��ȡ����������
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '��ȡ����������
    
    CH = IIf(Left(bh, 1) = "G", "��", "�㵥λ")
    T0 = "��ִ��" & bh & "��"
    t1 = "ͨ�������ʼ���ʽ���������Ϣ�������룬�����á�" & bt & "����Ϣ��"
    t2 = "���飬" & CH & "��������Ϊ���ϡ��л����񹲺͹�������Ϣ�����������ڶ�ʮ���涨���Ҿ���������" & Chr(13)
    t3 = "�������ݡ��л����񹲺͹�������Ϣ�����������ڶ�ʮ��������" & CH & "�����룬�Ҿֽ���"
    t4 = "ǰ��������𸴡�" & Chr(13) & "�����ش˸�֪��"
    t5 = "����ͳ�ƾ�ͳ�����Ϲ�������" & Chr(13)
    
    dt1 = Year(dt) & "��" & Month(dt) & "��" & Day(dt) & "��"
    
    td = Date
    fd = DateAdd("d", 21, td)
    dt2 = Year(fd) & "��" & Month(fd) & "��" & Day(fd) & "��"
    dt3 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
    Documents.Open ("D:\��Ϣ����\�Ǽǻ�ִ.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "��"
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="����" & dt1 & "��" & CH & t1 & t2 & t3 & dt2 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="��ִ" & bh, FileFormat:=wdFormatDocument    '�����ļ�
    
    Exit Do
    Loop

End Sub

'-------------------------------------------------------------------------
'��������Ĺ��ܣ���ȡ�������Ϣ���Զ����ɲ��ֹ�����֪�鲢���浽ָ��Ŀ¼
'���ߣ������
'���ڣ�2008��5��8��
'-------------------------------------------------------------------------

Sub G03_����()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Set MyDOC = Application.ActiveWindow.Document  'ȷ����ĵ�Ϊ�����ĵ�
    If MyDOC.Tables.Count > 0 Then  '�ж��ĵ����Ƿ��б��
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        If TN = "����ͳ�ƾ�������Ϣ���������" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
            '�ж��Ƿ��Ѿ��Ǽ�
            If Len(bh) < 10 Then
                MsgBox ("�������δ�Ǽǣ����ȵǼǣ�")
                mycheck = False
            Else
                '�ж��Ƿ��Ѿ����ɹ��Ǽǻ�ִ
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\��Ϣ����"
                        .FileName = "����" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("�Ѿ����ɹ��Ǽǻ�ִ�������ٴ����ɣ�")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("��������Ϣ�������������ִ�б������")
            mycheck = False
        End If
    Else
        MsgBox ("�ĵ���û�б�񣬲���ִ�б������")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '��ȡ���������
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '��ȡ��������λ����
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '��ȡ����������
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '��ȡ����������
    
    CH = IIf(Left(bh, 1) = "G", "��", "�㵥λ")
    T0 = "�����" & bh & "��"
    t1 = "�Ҿ�������" & CH & "�����������Ϣ�������룬��������Ǽǻ�ִ����" & bh & "�š�" & Chr(13)
    t2 = "�������飬" & CH & "�����ȡ����Ϣ���ڲ��ֹ�����Χ�����ݡ��л����񹲺͹�������Ϣ�����������ڶ�ʮ�����涨���Ҿֽ��Ե����ʼ���ʽ�ṩ���Թ������ֵ�������Ϣ��" & Chr(13)
    t3 = "����" & CH & "�����ȡ��������Ϣ�У��в����������ڣ�" & Chr(11) & "������ �����ͳ��������Ҫ�����ټӹ�������" & Chr(13)
    t3 = t3 & "������ ���һ���" & Chr(13) & "������ ��ҵ���ܻ��߹������ܵ�����ҵ���ܱ�й¶��������Ϣ" & Chr(13)
    t3 = t3 & "������ ���ڸ�����˽���߹������ܵ��¶Ը�����˽Ȩ��ɲ����ֺ���������Ϣ" & Chr(13) & "������ ���ɡ�����涨���蹫�����������Ρ�" & Chr(13)
    t4 = "�������ݡ��л����񹲺͹�������Ϣ������������ʮ����������" & CH & "�����ȡ�Ĳ�����Ϣ���Ҿֲ��蹫����" & Chr(13) & "�����ش˸�֪��"
    t5 = "����ͳ�ƾ�ͳ�����Ϲ�������" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "��" & sl_m & "��" & sl_d & "��"
    dt2 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) + 21 & "��"
    dt3 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
    Documents.Open ("D:\��Ϣ����\G02_���ֹ�����֪��.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "��"
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="����" & dt1 & "��" & t1 & t2 & t3 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="����" & bh, FileFormat:=wdFormatDocument    '�����ļ�
    
    Exit Do
    Loop

End Sub

'-------------------------------------------------------------------------------
'��������Ĺ��ܣ���ȡ�������Ϣ���Զ�����������Ϣ���蹫����֪�鲢���浽ָ��Ŀ¼
'���ߣ������
'���ڣ�2008��5��8��
'-------------------------------------------------------------------------------

Sub G04_����()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Set MyDOC = Application.ActiveWindow.Document  'ȷ����ĵ�Ϊ�����ĵ�
    If MyDOC.Tables.Count > 0 Then  '�ж��ĵ����Ƿ��б��
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        If TN = "����ͳ�ƾ�������Ϣ���������" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
            '�ж��Ƿ��Ѿ��Ǽ�
            If Len(bh) < 10 Then
                MsgBox ("�������δ�Ǽǣ����ȵǼǣ�")
                mycheck = False
            Else
                '�ж��Ƿ��Ѿ����ɹ��Ǽǻ�ִ
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\��Ϣ����"
                        .FileName = "����" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("�Ѿ����ɹ����蹫����֪�飬�����ٴ����ɣ�")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("��������Ϣ�������������ִ�б������")
            mycheck = False
        End If
    Else
        MsgBox ("�ĵ���û�б�񣬲���ִ�б������")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '��ȡ���������
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '��ȡ��������λ����
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '��ȡ����������
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '��ȡ����������
    
    CH = IIf(Left(bh, 1) = "G", "��", "�㵥λ")
    T0 = "�����" & bh & "��"
    t1 = "�Ҿ�������" & CH & "�����������Ϣ�������룬��������Ǽǻ�ִ����" & bh & "�š�"
    t2 = "���飬" & CH & "�����ȡ����Ϣ���ڣ�" & Chr(13)
    t3 = "������ �����ͳ��������Ҫ�����ټӹ�������" & Chr(13)
    t3 = t3 & "������ ���һ���" & Chr(13) & "������ ��ҵ���ܻ��߹������ܵ�����ҵ���ܱ�й¶��������Ϣ" & Chr(13)
    t3 = t3 & "������ ���ڸ�����˽���߹������ܵ��¶Ը�����˽Ȩ��ɲ����ֺ���������Ϣ" & Chr(13) & "������ ���ɡ�����涨���蹫�����������Ρ�" & Chr(13)
    t4 = "�������ݡ��л����񹲺͹�������Ϣ�����������ڶ�ʮһ���ڶ������" & CH & "�����ȡ����Ϣ���Ҿֲ��蹫����" & Chr(13) & "�����ش˸�֪��"
    t5 = "����ͳ�ƾ�ͳ�����Ϲ�������" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "��" & sl_m & "��" & sl_d & "��"
    dt2 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) + 21 & "��"
    dt3 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
    Documents.Open ("D:\��Ϣ����\G03_���蹫����֪��.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "��"
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="����" & dt1 & "��" & t1 & t2 & t3 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="����" & bh, FileFormat:=wdFormatDocument    '�����ļ�
    
    Exit Do
    Loop

End Sub

'-----------------------------------------------------------------------------
'��������Ĺ��ܣ���ȡ�������Ϣ���Զ�����������Ϣ�����ڸ�֪�鲢���浽ָ��Ŀ¼
'���ߣ������
'���ڣ�2008��5��8��
'-----------------------------------------------------------------------------

Sub G05_�����()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Set MyDOC = Application.ActiveWindow.Document  'ȷ����ĵ�Ϊ�����ĵ�
    If MyDOC.Tables.Count > 0 Then  '�ж��ĵ����Ƿ��б��
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        If TN = "����ͳ�ƾ�������Ϣ���������" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
            '�ж��Ƿ��Ѿ��Ǽ�
            If Len(bh) < 10 Then
                MsgBox ("�������δ�Ǽǣ����ȵǼǣ�")
                mycheck = False
            Else
                '�ж��Ƿ��Ѿ����ɹ��Ǽǻ�ִ
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\��Ϣ����"
                        .FileName = "�����" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("�Ѿ����ɹ���Ϣ�����ڸ�֪�飬�����ٴ����ɣ�")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("��������Ϣ�������������ִ�б������")
            mycheck = False
        End If
    Else
        MsgBox ("�ĵ���û�б�񣬲���ִ�б������")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '��ȡ���������
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '��ȡ��������λ����
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '��ȡ����������
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '��ȡ����������
    
    CH = IIf(Left(bh, 1) = "G", "��", "�㵥λ")
    T0 = "������" & bh & "��"
    t1 = "�Ҿ�������" & CH & "�����������Ϣ�������룬��������Ǽǻ�ִ����" & bh & "�š�" & Chr(13)
    t2 = "�������飬" & CH & "�����ȡ����Ϣ�����ڡ�" & Chr(13)
'    T3 = "�������ݡ��л����񹲺͹�������Ϣ�����������ڶ�ʮ��������" & CH & "�����룬�Ҿֽ���"
    t4 = "�����ش˸�֪��"
    t5 = "����ͳ�ƾ�ͳ�����Ϲ�������" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "��" & sl_m & "��" & sl_d & "��"
    dt2 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) + 21 & "��"
    dt3 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
    Documents.Open ("D:\��Ϣ����\G04_��Ϣ�����ڸ�֪��.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "��"
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="����" & dt1 & "��" & t1 & t2 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="�����" & bh, FileFormat:=wdFormatDocument    '�����ļ�
    
    Exit Do
    Loop

End Sub

'-------------------------------------------------------------------------------------
'��������Ĺ��ܣ���ȡ�������Ϣ���Զ����ɷǱ�����������Ϣ������֪�鲢���浽ָ��Ŀ¼
'���ߣ������
'���ڣ�2008��5��8��
'-------------------------------------------------------------------------------------

Sub G06_�Ǹ�()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Set MyDOC = Application.ActiveWindow.Document  'ȷ����ĵ�Ϊ�����ĵ�
    If MyDOC.Tables.Count > 0 Then  '�ж��ĵ����Ƿ��б��
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        If TN = "����ͳ�ƾ�������Ϣ���������" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
            '�ж��Ƿ��Ѿ��Ǽ�
            If Len(bh) < 10 Then
                MsgBox ("�������δ�Ǽǣ����ȵǼǣ�")
                mycheck = False
            Else
                '�ж��Ƿ��Ѿ����ɹ��Ǽǻ�ִ
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\��Ϣ����"
                        .FileName = "�Ǹ�" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("�Ѿ����ɹ��Ǳ�������Ϣ֪�飬�����ٴ����ɣ�")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("��������Ϣ�������������ִ�б������")
            mycheck = False
        End If
    Else
        MsgBox ("�ĵ���û�б�񣬲���ִ�б������")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '��ȡ���������
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '��ȡ��������λ����
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '��ȡ����������
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '��ȡ����������
    
    CH = IIf(Left(bh, 1) = "G", "��", "�㵥λ")
    T0 = "�Ǹ��" & bh & "��"
    t1 = "�Ҿ�������" & CH & "�����������Ϣ�������룬��������Ǽǻ�ִ����" & bh & "�š�" & Chr(13)
    t2 = "�������飬" & CH & "�����ȡ����Ϣ�����ڱ����ص����շ�Χ��������____������ѯ����ϵ��ʽΪ____��" & Chr(13)
'    T3 = "�������ݡ��л����񹲺͹�������Ϣ�����������ڶ�ʮ��������" & CH & "�����룬�Ҿֽ���"
    t4 = "�����ش˸�֪��"
    t5 = "����ͳ�ƾ�ͳ�����Ϲ�������" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "��" & sl_m & "��" & sl_d & "��"
    dt2 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) + 21 & "��"
    dt3 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
    Documents.Open ("D:\��Ϣ����\G05_�Ǳ�����������Ϣ��֪��.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "��"
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="����" & dt1 & "��" & t1 & t2 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="�Ǹ�" & bh, FileFormat:=wdFormatDocument    '�����ļ�
    
    Exit Do
    Loop

End Sub

'-------------------------------------------------------------------------
'��������Ĺ��ܣ���ȡ�������Ϣ���Զ����ɲ�������֪ͨ�鲢���浽ָ��Ŀ¼
'���ߣ������
'���ڣ�2008��5��8��
'-------------------------------------------------------------------------

Sub G07_��ͨ()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Set MyDOC = Application.ActiveWindow.Document  'ȷ����ĵ�Ϊ�����ĵ�
    If MyDOC.Tables.Count > 0 Then  '�ж��ĵ����Ƿ��б��
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        If TN = "����ͳ�ƾ�������Ϣ���������" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
            '�ж��Ƿ��Ѿ��Ǽ�
            If Len(bh) < 10 Then
                MsgBox ("�������δ�Ǽǣ����ȵǼǣ�")
                mycheck = False
            Else
                '�ж��Ƿ��Ѿ����ɹ��Ǽǻ�ִ
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\��Ϣ����"
                        .FileName = "��ͨ" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("�Ѿ����ɹ���������֪ͨ�飬�����ٴ����ɣ�")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("��������Ϣ�������������ִ�б������")
            mycheck = False
        End If
    Else
        MsgBox ("�ĵ���û�б�񣬲���ִ�б������")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '��ȡ���������
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '��ȡ��������λ����
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '��ȡ����������
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '��ȡ����������
    
    CH = IIf(Left(bh, 1) = "G", "��", "�㵥λ")
    T0 = "��ͨ��" & bh & "��"
    t1 = "�Ҿ�������" & CH & "�����������Ϣ�������룬��������Ǽǻ�ִ����" & bh & "�š�" & Chr(13)
    t2 = "�������飬" & CH & "�����ȡ����Ϣ���ݲ���ȷ���Ҿ����Ը��ݴ�����ȷ�������������Ϣ������ġ��������������������������롣" & Chr(13)
'    T3 = "�������ݡ��л����񹲺͹�������Ϣ�����������ڶ�ʮ��������" & CH & "�����룬�Ҿֽ���"
    t4 = "�����ش˸�֪��"
    t5 = "����ͳ�ƾ�ͳ�����Ϲ�������" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "��" & sl_m & "��" & sl_d & "��"
    dt2 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) + 21 & "��"
    dt3 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
    Documents.Open ("D:\��Ϣ����\G06_��������֪ͨ��.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "��"
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="����" & dt1 & "��" & t1 & t2 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="��ͨ" & bh, FileFormat:=wdFormatDocument    '�����ļ�
    
    Exit Do
    Loop

End Sub

'-------------------------------------------------------------------------
'��������Ĺ��ܣ���ȡ�������Ϣ���Զ����ɲ�������֪ͨ�鲢���浽ָ��Ŀ¼
'���ߣ������
'���ڣ�2008��5��8��
'-------------------------------------------------------------------------

Sub G08_��֪()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Set MyDOC = Application.ActiveWindow.Document  'ȷ����ĵ�Ϊ�����ĵ�
    If MyDOC.Tables.Count > 0 Then  '�ж��ĵ����Ƿ��б��
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        If TN = "����ͳ�ƾ�������Ϣ���������" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
            '�ж��Ƿ��Ѿ��Ǽ�
            If Len(bh) < 10 Then
                MsgBox ("�������δ�Ǽǣ����ȵǼǣ�")
                mycheck = False
            Else
                '�ж��Ƿ��Ѿ����ɹ��Ǽǻ�ִ
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\��Ϣ����"
                        .FileName = "��֪" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("�Ѿ����ɹ�������֪�飬�����ٴ����ɣ�")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("��������Ϣ�������������ִ�б������")
            mycheck = False
        End If
    Else
        MsgBox ("�ĵ���û�б�񣬲���ִ�б������")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '��ȡ���������
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '��ȡ��������λ����
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '��ȡ����������
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '��ȡ����������
    
    CH = IIf(Left(bh, 1) = "G", "��", "�㵥λ")
    T0 = "��֪��" & bh & "��"
    t1 = "�Ҿ�������" & CH & "�����������Ϣ�������룬��������Ǽǻ�ִ����" & bh & "�š�" & Chr(13)
    t2 = "�������飬" & CH & "�����ȡ����Ϣ���ڹ�����Χ�����ݡ��л����񹲺͹�������Ϣ�����������ڶ�ʮһ����һ�"
    t3 = "�Ҿֽ��Ե����ʼ���ʽ�ṩ�������������Ϣ��" & Chr(13)
    t4 = "�����ش˸�֪��"
    t5 = "����ͳ�ƾ�ͳ�����Ϲ�������" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "��" & sl_m & "��" & sl_d & "��"
    dt2 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) + 21 & "��"
    dt3 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
    Documents.Open ("D:\��Ϣ����\G01_������֪��.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:=xm & "��"
        mt.Cell(5, 1).Select
        Selection.TypeText Text:="����" & dt1 & "��" & t1 & t2 & t3 & t4
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="��֪" & bh, FileFormat:=wdFormatDocument    '�����ļ�
    
    Exit Do
    Loop

End Sub


'-------------------------------------------------------------------------
'��������Ĺ��ܣ���ȡ�������Ϣ���Զ����ɲ�������֪ͨ�鲢���浽ָ��Ŀ¼
'���ߣ������
'���ڣ�2008��5��8��
'-------------------------------------------------------------------------

Sub G09_Э��()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Set MyDOC = Application.ActiveWindow.Document  'ȷ����ĵ�Ϊ�����ĵ�
    If MyDOC.Tables.Count > 0 Then  '�ж��ĵ����Ƿ��б��
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        If TN = "����ͳ�ƾ�������Ϣ���������" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
            '�ж��Ƿ��Ѿ��Ǽ�
            If Len(bh) < 10 Then
                MsgBox ("�������δ�Ǽǣ����ȵǼǣ�")
                mycheck = False
            Else
                '�ж��Ƿ��Ѿ����ɹ��Ǽǻ�ִ
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\��Ϣ����"
                        .FileName = "Э��" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("�Ѿ����ɹ�Э��֪ͨ�飬�����ٴ����ɣ�")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("��������Ϣ�������������ִ�б������")
            mycheck = False
        End If
    Else
        MsgBox ("�ĵ���û�б�񣬲���ִ�б������")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '��ȡ���������
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '��ȡ��������λ����
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '��ȡ����������
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '��ȡ����������
        yt = Left(mytable.Cell(17, 3).Range.Text, Len(mytable.Cell(17, 3).Range.Text) - 2) '��ȡ����;
        dh = Left(mytable.Cell(9, 3).Range.Text, Len(mytable.Cell(9, 3).Range.Text) - 2) '��ȡ����ϵ�绰
        em = Left(mytable.Cell(11, 3).Range.Text, Len(mytable.Cell(11, 3).Range.Text) - 2) '��ȡ��email
    
    t3 = "�����ֽ��Ҿ��ѵǼ�����ĵ�" & bh
    t3 = t3 & "��������Ϣ��������ת������λ����Э����������Ϊʮ�������ա�"
    t3 = t3 & "���ڰ�����Ϻ󾡿콫�������������ǣ���������ͳһ�������ˡ�лл��" & Chr(13)
    t4 = "�����ش�֪ͨ��"
    t5 = "����ͳ�ƾ�ͳ�����Ϲ�������" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "��" & sl_m & "��" & sl_d & "��"
    dt2 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) + 21 & "��"
    dt3 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
    
    Documents.Open ("D:\��Ϣ����\G07_Э��֪ͨ.doc")
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
    
    ActiveDocument.SaveAs FileName:="Э��" & bh, FileFormat:=wdFormatDocument    '�����ļ�
    
    Exit Do
    Loop

End Sub

'-------------------------------------------------------------------------
'��������Ĺ��ܣ���ȡ�������Ϣ���Զ����ɲ�������֪ͨ�鲢���浽ָ��Ŀ¼
'���ߣ������
'���ڣ�2008��5��8��
'-------------------------------------------------------------------------

Sub G10_Э��()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Set MyDOC = Application.ActiveWindow.Document  'ȷ����ĵ�Ϊ�����ĵ�
    If MyDOC.Tables.Count > 0 Then  '�ж��ĵ����Ƿ��б��
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        If TN = "����ͳ�ƾ�������Ϣ���������" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
            '�ж��Ƿ��Ѿ��Ǽ�
            If Len(bh) < 10 Then
                MsgBox ("�������δ�Ǽǣ����ȵǼǣ�")
                mycheck = False
            Else
                '�ж��Ƿ��Ѿ����ɹ��Ǽǻ�ִ
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\��Ϣ����"
                        .FileName = "Э��" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("�Ѿ����ɹ�Э��֪ͨ�飬�����ٴ����ɣ�")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("��������Ϣ�������������ִ�б������")
            mycheck = False
        End If
    Else
        MsgBox ("�ĵ���û�б�񣬲���ִ�б������")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '��ȡ���������
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '��ȡ��������λ����
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '��ȡ����������
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '��ȡ����������
        yt = Left(mytable.Cell(17, 3).Range.Text, Len(mytable.Cell(17, 3).Range.Text) - 2) '��ȡ����;
        dh = Left(mytable.Cell(9, 3).Range.Text, Len(mytable.Cell(9, 3).Range.Text) - 2) '��ȡ����ϵ�绰
        em = Left(mytable.Cell(11, 3).Range.Text, Len(mytable.Cell(11, 3).Range.Text) - 2) '��ȡ��email
    
    t3 = "�����ֽ��Ҿ��ѵǼ�����ĵ�" & bh
    t3 = t3 & "��������Ϣ��������ת������λ����Э���й�˾����λ��������Ϊʮ�������ա�"
    t3 = t3 & "�������쵥λ�ڰ�����Ϻ󾡿콫�������������ǣ���������ͳһ�������ˡ�лл��" & Chr(13)
    t4 = "�����ش�֪ͨ��"
    t5 = "����ͳ�ƾ�ͳ�����Ϲ�������" & Chr(13)
    
    sl_y = Left(Right(bh, Len(bh) - 1), 4)
    sl_m = Left(Right(bh, Len(bh) - 5), 2)
    sl_m = IIf(Left(sl_m, 1) = "0", Right(sl_m, 1), sl_m)
    sl_d = Left(Right(bh, Len(bh) - 7), 2)
    sl_d = IIf(Left(sl_d, 1) = "0", Right(sl_d, 1), sl_d)

    dt1 = sl_y & "��" & sl_m & "��" & sl_d & "��"
    dt2 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) + 21 & "��"
    dt3 = Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
    
    Documents.Open ("D:\��Ϣ����\G08_Э��֪ͨ.doc")
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
    
    ActiveDocument.SaveAs FileName:="Э��" & bh, FileFormat:=wdFormatDocument    '�����ļ�
    
    Exit Do
    Loop

End Sub


'----------------------------------------------------------------------------
'��������Ĺ��ܣ��Զ�����ͬ������������б�š����桢�Ǽ�
'���ߣ������
'���ڣ�2008��5��5��
'----------------------------------------------------------------------------
Sub E01_�Ǽ�()

    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Set MyDOC = Application.ActiveWindow.Document  'ȷ����ĵ�Ϊ�����ĵ�
    
    If MyDOC.Tables.Count = 0 Then  '�ж��ĵ����Ƿ��б��
        MsgBox ("�ĵ���û�б�񣬲���ִ�б������")
        mycheck = False
    Else
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        MsgBox TN
        If TN <> "Application Form for the Publication of Government Information of NBS" Then
            MsgBox ("��������Ϣ�������������ִ�б������")
            mycheck = False
        Else
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
            lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '��ȡ�������
            dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '��ȡ����������
            '�ж��Ƿ��Ѿ��Ǽ�
            If Len(bh) < 10 Then
                If IsDate(dt) Then
                    mycheck = True
                Else
                    MsgBox "�������ڸ�ʽ���������޸ģ���ִ�б����" & Chr(13) & "��ȷ�ĸ�ʽΪ��2008-5-1��2008-05-01��2008��5��1��"
                    mycheck = False
                End If
            Else
                MsgBox ("��������ѵǼǣ������ظ��Ǽǣ�")
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

'�ڹ���Ŀ¼�в����Ƿ��е��������ļ������û�У���Ŵ�1��ʼ��������ԭ���ı�Ż������������
Set fs = Application.FileSearch
With fs
    .LookIn = "D:\��Ϣ����"
    .FileName = "����" & N
    
    If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
        FN = N & 1
    Else
        FN = N & .FoundFiles.Count + 1
    End If
End With

    If Len(ActiveDocument.Tables(1).Cell(2, 2).Range.Text) < Len(FN) Then
    ActiveDocument.Tables(1).Cell(2, 2).Select
    Selection.TypeText Text:=FN
    ActiveDocument.SaveAs FileName:="����" & FN, FileFormat:=wdFormatDocument   '�����ļ�
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

    Documents.Open FileName:="D:\��Ϣ����\�ǼǱ�3.doc"
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
'    MsgBox ("�ѳɹ��Ǽǣ�")
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

    Documents.Open FileName:="D:\��Ϣ����\�ǼǱ�4.doc"
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
'    MsgBox ("�ѳɹ��Ǽǣ�")
    End If
    
    Exit Do
    Loop
    
End Sub

'--------------------------------------------------------------
'��������Ĺ��ܣ���ȡ�������Ϣ���Զ����ɻ�ִ�����浽ָ��Ŀ¼
'���ߣ������
'���ڣ�2008��5��6��
'--------------------------------------------------------------

Sub E02_��ִ()
    
    Dim mycheck As Boolean
    Dim MyDOC, mytable
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Set MyDOC = Application.ActiveWindow.Document  'ȷ����ĵ�Ϊ�����ĵ�
    If MyDOC.Tables.Count > 0 Then  '�ж��ĵ����Ƿ��б��
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        MsgBox TN
        If TN = "Application Form for the Publication of Government Information of NBS" Then
            bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
            '�ж��Ƿ��Ѿ��Ǽ�
            If Len(bh) < 10 Then
                MsgBox ("�������δ�Ǽǣ����ȵǼǣ�")
                mycheck = False
            Else
                '�ж��Ƿ��Ѿ����ɹ��Ǽǻ�ִ
                Set fs = Application.FileSearch
                With fs
                     .LookIn = "D:\��Ϣ����"
                        .FileName = "��ִ" & bh
                         If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) = 0 Then
                            mycheck = True
                         Else
                            MsgBox ("�Ѿ����ɹ��Ǽǻ�ִ�������ٴ����ɣ�")
                            mycheck = False
                         End If
                End With
            End If
        Else
            MsgBox ("��������Ϣ�������������ִ�б������")
            mycheck = False
        End If
    Else
        MsgBox ("�ĵ���û�б�񣬲���ִ�б������")
        mycheck = False
    End If
    
    Do While mycheck = True
        Set mytable = ActiveDocument.Tables(1)
        TN = Left(mytable.Cell(1, 1).Range.Text, Len(mytable.Cell(1, 1).Range.Text) - 2) '��ȡ���������
        lb = Left(mytable.Cell(2, 3).Range.Text, Len(mytable.Cell(2, 3).Range.Text) - 2) '��ȡ���������
        bh = Left(mytable.Cell(2, 2).Range.Text, Len(mytable.Cell(2, 2).Range.Text) - 2) '��ȡ�����
        xm = Left(mytable.Cell(5, 3).Range.Text, Len(mytable.Cell(5, 3).Range.Text) - 2) '��ȡ��������λ����
        dt = Left(mytable.Cell(14, 3).Range.Text, Len(mytable.Cell(14, 3).Range.Text) - 2) '��ȡ����������
        bt = Left(mytable.Cell(16, 3).Range.Text, Len(mytable.Cell(16, 3).Range.Text) - 2) '��ȡ����������
    
'    CH = IIf(Left(bh, 1) = "G", "��", "�㵥λ")
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
    
    Documents.Open ("D:\��Ϣ����\E01_Feedback.doc")
    Set mt = ActiveDocument.Tables(1)
        mt.Cell(3, 1).Select
        Selection.TypeText Text:=T0
        mt.Cell(4, 1).Select
        Selection.TypeText Text:="Mr. /Mrs. " & xm & ","
        mt.Cell(5, 1).Select
        Selection.TypeText Text:=t1 & dt2 & " by NBS." & Chr(13) & t6
        mt.Cell(6, 2).Select
        Selection.TypeText Text:=t5 & dt3
    
    ActiveDocument.SaveAs FileName:="��ִ" & bh, FileFormat:=wdFormatDocument    '�����ļ�
    
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
        
        MsgBox "������" & x1 & Chr(13) & "��λ��" & x2 & " " & Chr(13) & "ְ��" & x3 & Chr(13) & "�ֻ���" & x4 & Chr(13) & "ֱ������" & x5 & Chr(13) & "���ʣ�" & x6 & Chr(13)
        
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
'��������Ĺ��ܣ��Զ���������Ų����浽ָ��Ŀ¼
'���ߣ������
'���ڣ�2008��5��7��
'---------------------------------------------------

Sub B01_����()

    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    fl = "NBSTEL.TXT"    ' ��Ҫת�����ļ���
    DL = Chr(9)       ' �ָ���
    TXT = ""            ' ����һ�����ַ�������
    NM = InputBox("��������Ҫ��ѯ�����ݣ�")
    
    Open fl For Input As #1     ' �������ļ���
    MsgBox LOF(1)
    Do While Not EOF(1)         ' ѭ�����ļ�β��
        Line Input #1, LN       ' �ѵ�һ���ַ�����ֵ������LN
        If InStr(1, LN, NM, 1) Then
'            MsgBox LN
        L = Split(LN, delimiter:=DL) ' ���ָ������ַ�����ֵ������L
        MsgBox "������" & L(0) & Chr(13) & "��λ��" & L(1) & " " & Chr(13) & "ְ��" & L(2) & Chr(13) & "�ֻ���" & L(3) & Chr(13) & "ֱ����" & L(4) & Chr(13) & "���ʣ�" & L(5) & Chr(13), , "��ѯ���"
        End If
        
'        L = Split(LN, delimiter:=DL) ' ���ָ������ַ�����ֵ������L
'        For i = 0 To UBound(L)
'        txt = txt & Chr(9) & Trim(L(i))
'        Next i
'        txt = Right(txt, Len(txt) - 1)
'        Selection.TypeText Text:=txt & Chr(13)
    Loop        ' ѭ��
            L = Loc(1)
            MsgBox Loc(1)

    Close #1    ' �ر��ļ�

    Open fl For Output As #1     ' �������ļ���

Seek #1, L
Write #1, Chr(13) & "This is a test!" & Chr(13)
 Close #1    ' �ر��ļ�

End Sub


'---------------------------------------------------
'��������Ĺ��ܣ��Զ���������Ų����浽ָ��Ŀ¼
'���ߣ������
'���ڣ�2008��5��27��
'---------------------------------------------------

Sub �Զ���ƴ��()
    ChangeFileOpenDirectory "D:\��Ϣ����"      'ȷ������Ŀ¼
    Dim Z() As Variant, P() As Variant, Q() As Variant
    
    fl = "pinyin.doc"    ' ƴ�����ձ��ļ���
    DL = Chr(9)       ' �ָ���
    txt1 = ""            ' ����һ�����ַ�������
    txt2 = ""
    Do While True
    NM = InputBox("��������Ҫ��ѯ�����ݣ�")
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

    
    Documents.Open FileName:=fl, Visible:=False   ' ��ƴ�����ձ��ļ�
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

Sub �ļ��Ի���()
    Dim fd As FileDialog    '����һ���ļ��Ի������
    Set fd = Application.FileDialog(msoFileDialogFilePicker)    '�����ļ��Ի������
    Dim fs As Variant
    With fd
        .AllowMultiSelect = False   'ֻ����ѡ��һ���ļ�
        If .Show = -1 Then  '��show��������ʾ�ļ�ѡ��Ի���
            For Each fs In .SelectedItems   '����FileDialogSelectedItems����ÿ����Ա
                MsgBox "ѡ����ļ�·��Ϊ: " & fs
            Next fs
        Else    '�û�����ȡ����
        End If
    End With
    Set fd = Nothing    '�ͷŶ���Ķ���

End Sub

Sub ���Ҵ���()
     Do While True
    NM = InputBox("��������Ҫ��ѯ�����ݣ�")
    If Len(NM) > 0 Then
    
    Exit Do
    End If
    Loop
    Set MyDoc1 = ActiveDocument
        
        fl = "C:\EasyQuery\NBSTEL.doc"
        Documents.Open FileName:=fl, Visible:=False   ' ���ļ�
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
        MsgBox "�ҵ� " & i & " ��", , "��ʾ"
    Else
        MsgBox "δ�ҵ���", , "��ʾ"
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
 '       MsgBox "�ڵ� " & n & " ���ҵ�һ��"
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
        
        MsgBox "���ҵ� " & i & " ����" & "�����Ķ���ֱ��ǣ�" & st, , "��ʾ"
    Else
        MsgBox "δ�ҵ���", , "��ʾ"
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
    fc = "C:\EasyQuery\��������ļ�.doc"
    Documents.Open FileName:=fc, Visible:=True
    Documents(fc).Activate
    Set my_doc = ActiveDocument
'    If ActiveDocument.Tables.Count > 0 Then
'        Set mytab = ActiveDocument.Tables(1)
'        cn = mytab.Columns.Count
'        rn = mytab.Rows.Count
'        MsgBox "���������" & cn & "  " & "���������" & rn
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
    fname = "C:\EasyQuery\pinyin.doc"    ' ƴ�����ձ��ļ���
    txt1 = ""            ' ����һ�����ַ�������
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
    Doc_Open (fname) ' �򿪻򼤻�ƴ�����ձ��ļ�
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
Function Doc_Open(fname) '�ж��ļ��Ƿ��Ѿ��򿪣�����Ѵ򿪣��򼤻���򣬾ʹ�
    For Each doc In Documents
        If doc.name = fname Then Found = True
    Next doc
    If Found <> True Then
        Documents.Open FileName:=fname, Visible:=False  ' �򿪵绰�����ļ�
    Else
        Documents(fname).Activate
    End If
End Function
Function Doc_Close(fname) '�ж��ļ��Ƿ��Ѿ��򿪣�����Ѵ򿪣���ر�
    For Each doc In Documents
        If doc.name = fname Then Found = True
    Next doc
    If Found = True Then Documents(fname).Close
End Function
Sub A01_�ļ�·��()
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
'        Show_Info ("û���ҵ���Ҫ��ѯ�����ݣ�����������")
        MsgBox "û���ҵ���Ҫ��ѯ�����ݣ�����������"
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
'        Show_Info ("û���ҵ���Ҫ��ѯ�����ݣ�����������")
        MsgBox "û���ҵ���Ҫ��ѯ�����ݣ�����������"
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
        MsgBox "��Ҫ�ҵ��ǣ�" & Name1 & Chr(13) & "    ����Ϊ��" & Code1 & Chr(13) & "    �����ڣ�" & Name2 & Name3
    Else
        MsgBox "û���ҵ���"
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

Sub A01_����Ƿ�������Word()
'Documents("NBSTEL.doc").Close
If Tasks.Exists(name:="Microsoft Word") = True Then
    Set myobject = GetObject("", "Word.Application")
'    MsgBox myobject.Name
    Application.Quit SaveChanges:=wdSaveChanges
    Set myobject = Nothing
End If
'Set Word = CreateObject("word.basic")
'Set msword = CreateObject("word.application")
'Set mydoc = msword.Documents.Open("C:\NBS�绰��ѯ\NBSTEL.doc", PasswordDocument:="nbsdhg", WritePasswordDocument:="nbsdhg")
'msword.Visible = True

End Sub

Sub GetWord()
    Dim MyWord As Object    '���ڴ��'Microsoft Word ���õı�����
    Dim NoWord As Boolean    '��������ͷŵı�ǡ�
    On Error Resume Next    '�ӳٴ��󲶻�
    Set MyWord = GetObject(, "Word.Application")
    If Err.Number <> 0 Then NoWord = True
    Err.Clear    '�������������Ҫ��� Err ����
    If NoWord <> True Then
'        MyWord.Application.quit
        MyWord.Visible = False
        MsgBox "Hide Word Successfully!"
    Else
        Set MyWord = CreateObject("Word.Application")
        MyWord.Documents.Open ("D:\00 F2008\����λ�绰\�������ģ�ȫ��.doc")
        MyWord.Visible = False
        MsgBox "�ɹ�����Word����" & MyWord.ActiveDocument.name
    End If
    MsgBox "Will Quit!"
    MyWord.Quit
    Unload Form1
End Sub

Sub A01_CSV() '�����е����ʼ���ַ�ļ�¼������CSV�ļ�
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set A = fs.CreateTextFile("C:\EasyQuery\Add.csv", True)
    j = ","
    A.WriteLine ("����,����,ְ��,ҵ��绰,�ƶ��绰,סլ�绰,�����ʼ���ַ,�칫�ص�")
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
    MsgBox "�����ɹ���"
    Documents("NBSTEL.DOC").Close
End Sub


'---------------------------------------------------
'���ܣ���ѯ�ֻ�����
'���ߣ������
'���ڣ�2008��6��10��
'---------------------------------------------------
Sub A01_Mobile() '�ֻ����ز�ѯ
    Dim tel As String, fname As String
    FN = "130 131 132 133 134 135 136 137 138 139 159"
    ChangeFileOpenDirectory "C:\EasyQuery" 'ȷ������Ŀ¼
    Do While True
        tel = InputBox("��������Ҫ��ѯ�ֻ����룺" & Chr(13) & "����������ǰ7λ��")
        If Len(tel) > 0 Then
            If Len(tel) > 6 Then
                If Asc(Left(tel, 1)) = 49 Then
                    If InStr(1, FN, Left(tel, 3), 1) > 0 Then
                        chk = True
                        Exit Do
                    Else
                        MsgBox "���ݿ���û��" & Left(tel, 3) & "��ͷ�ĺ���!"
                        chk = False
                        Exit Do
                    End If
                Else
                    MsgBox "���벻��ȷ���ֻ��ű����ԡ�1����ͷ�����������룡"
                    chk = False
                End If
            Else
                MsgBox "�����������ֻ��ŵ�ǰ7λ���֣���ֻ������" & Len(tel) & "λ��"
                chk = False
            End If
        Else
            MsgBox "��û�����룬�����������ֻ��ŵ�ǰ7λ���֣�"
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
        MsgBox "����ҵĺ�����������Ϊ��" & fd(1) & "����" & fd(2)
    Else
         MsgBox "û���ҵ���Ҫ��ѯ�����ݣ�"
    End If
    Documents(fname).Close
    Exit Do
    Loop
End Sub

Option Base 1

Sub A01_Txt_Import()
    Dim PA() As Variant, pn() As Variant, LN As String, MyDOC As Document
    i = 0
    Open "C:\EasyQuery\Test.txt" For Input As #1    ' �������ļ���
    Do While Not EOF(1)    ' ѭ�����ļ�β��
        Line Input #1, LN
        i = i + 1
        ReDim Preserve PA(i)
        PA(i) = LN
    Loop
    Close #1    ' �ر��ļ�
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
    ҳ���
    For j = 1 To UBound(pn)
    Selection.TypeText Text:=pn(j)
    Next j
    Selection.TypeBackspace
    Doc_Close ("C:\EasyQuery\pinyin.doc")

End Sub

Sub A01_����׼��()
    Dim chk As Boolean '����һ�������������ж��ĵ����Ƿ��б�����
    Dim MyTab As Table '����һ�����������
    N = ActiveDocument.Tables.Count 'ȡ�ñ������
    If N = 0 Then
        MsgBox "�ĵ���û�б�񣬲���ִ�иú����" '��ʾ�û��ĵ���û�б��
        chk = False
    Else
        chk = True
    End If
    On Error Resume Next
    Do While chk = True
    For i = 1 To N
    ActiveDocument.Tables(i).Range.Select
    Set MyTab = Selection.Tables(1)
    With MyTab '���ñ������
        .Rows.Alignment = wdAlignRowRight '������ݾ��Ҷ���
        .TopPadding = CentimetersToPoints(0) '���õ�Ԫ���ϱ߾�Ϊ0
        .BottomPadding = CentimetersToPoints(0) '���õ�Ԫ���±߾�Ϊ0
        .LeftPadding = CentimetersToPoints(0) '���õ�Ԫ����߾�Ϊ0
        .RightPadding = CentimetersToPoints(0) '���õ�Ԫ���ұ߾�Ϊ0
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone '��������
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    End With
        Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter '��Ԫ�����ݴ�ֱ����

    With Selection.Borders(wdBorderTop) '���ö���
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Borders(wdBorderBottom) '���õ���
        .LineStyle = Options.DefaultBorderLineStyle
        .LineWidth = Options.DefaultBorderLineWidth
        .Color = Options.DefaultBorderColor
    End With
    With Selection.Font '��������
        .NameFarEast = "����"
        .NameAscii = "Times New Roman"
        .NameOther = "Times New Roman"
        .Size = 10.5
    End With
    With Selection.ParagraphFormat '���ö����ʽ
        .LeftIndent = CentimetersToPoints(0.1)
        .RightIndent = CentimetersToPoints(0.1)
        .LineSpacingRule = wdLineSpaceAtLeast
        .LineSpacing = 12
        .WordWrap = True
    End With
    MyTab.Rows(1).Select 'ѡ�����ĵ�һ����Ԫ��
    Selection.SelectRow 'ѡ����һ��
    Selection.Rows.Height = CentimetersToPoints(1#) '���õ�һ�е��и�Ϊ1����
    

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

Sub A01_��()
    Dim MyDOC As Document, MyDir As String
    'MyDir = "D:\��2008"
    'FN = "���ڷ��������ӳ" & Year(Date) & Chr(45)
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
        .NameFarEast = "����"
        .NameAscii = "����"
        .Size = 15
    End With
    MyDOC.Paragraphs(1).Range.Select
    Selection.TypeText Text:="���ڷ��������ӳ"
    MyDOC.Paragraphs(1).Range.Select
    With Selection.ParagraphFormat
        .SpaceBefore = 30
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
    End With
    With Selection.Font
        .NameFarEast = "��������"
        .Size = 36
        .Bold = True
        .Color = wdColorRed
    End With
    
    MyDOC.Paragraphs(2).Range.Select
    Selection.TypeText Text:=Year(Date) & "��� " & "�ڣ��ܵ�  �ڣ�"
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
    Selection.TypeText Text:="����ͳ�ƾ�ͳ�����Ϲ�������" & vbTab & Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
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
    Selection.TypeText Text:="����"
    MyDOC.Paragraphs(5).Range.Select
        With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(0.35)
        .CharacterUnitFirstLineIndent = 2
    End With
    With Selection.Font
        .NameFarEast = "����"
        .NameAscii = "����"
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
   ' MyDoc.SaveAs FileName:=FN, FileFormat:=wdFormatDocument    '�����ļ�

End Sub

Sub A01_ÿ������()
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
        .NameFarEast = "����"
        .NameAscii = "����"
        .Size = 15
    End With
    MyDOC.Paragraphs(1).Range.Select
    Selection.TypeText Text:="ÿ������"
    MyDOC.Paragraphs(1).Range.Select
    With Selection.ParagraphFormat
        .SpaceBefore = 60
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphCenter
    End With
    With Selection.Font
        .NameFarEast = "�����п�"
        .Size = 72
        .Bold = False
        .Color = wdColorRed
    End With
    
    MyDOC.Paragraphs(2).Range.Select
    Selection.TypeText Text:=Year(Date) & "��� " & "�ڣ��ܵ�  �ڣ�"
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
    Selection.TypeText Text:="����ͳ�ƾ�ͳ�����Ϲ�������" & vbTab & Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
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
    Selection.TypeText Text:="����"
    MyDOC.Paragraphs(5).Range.Select
        With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceSingle
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = CentimetersToPoints(0.35)
        .CharacterUnitFirstLineIndent = 2
    End With
    With Selection.Font
        .NameFarEast = "����"
        .NameAscii = "����"
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
   ' MyDoc.SaveAs FileName:=FN, FileFormat:=wdFormatDocument    '�����ļ�

End Sub
Sub Char_To_Array()
    Dim name As String
    Dim NA() As Variant
    name = "ŷ����Ծ"
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
        MsgBox "��ѡ����Ϊ��������ݣ�"
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

Sub ��ҳͼƬ����()
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

Sub A00_����ָ��Ŀ¼()

NM = "F:\00 F" & Year(Date)
NM1 = NM & "\01 OA�ļ�"
NM2 = NM & "\02 ��������"
NM3 = NM & "\03 �ͼ칤��"
NM4 = NM & "\04 �ҵ��ļ�"
NM5 = NM & "\05 �е���Ŀ"
NM6 = NM & "\06 ������Ŀ"
NM7 = NM & "\07 ��������"
NM8 = NM & "\08 PDF"
NM9 = NM & "\09 ����"
NM10 = NM & "\10  ������ժ"

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


Sub A01_���ǰ�ı��Ӵ�()

    Application.ScreenUpdating = False '�ر���Ļ����
    C = Chr(-24157) '���
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
    
    Selection.HomeKey Unit:=wdStory '����������ĵ���ʼ��
    Application.ScreenUpdating = True '�ָ���Ļ����

End Sub

Sub A01_���������ı��Ӵ�()

    Application.ScreenUpdating = False '�ر���Ļ����
    C = Chr(93) '������]
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
    
    Selection.HomeKey Unit:=wdStory '����������ĵ���ʼ��
    Application.ScreenUpdating = True '�ָ���Ļ����

End Sub

Sub A01_ð��ǰ�ı��Ӵ�()

    Application.ScreenUpdating = False '�ر���Ļ����

    C = Chr(-23622) 'ð�ţ�
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
    
    Selection.HomeKey Unit:=wdStory '����������ĵ���ʼ��
    Application.ScreenUpdating = True '�ָ���Ļ����

End Sub


Sub ����ҳ��()
'
' ���� 2009-2-3 �� �����: ¼��
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
    Selection.TypeText Text:="��"
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Selection.Copy
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.PasteAndFormat (wdPasteDefault)
    Selection.MoveLeft Unit:=wdCharacter, Count:=3, Extend:=wdExtend
    Selection.Font.Size = 12
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub

Sub A01_�������()  '����������
    Selection.ParagraphFormat.TabStops.ClearAll
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(14.07), Alignment:=wdAlignTabCenter, Leader:=wdTabLeaderSpaces
    Selection.TypeText Text:=Chr(9) & "ͳ�����Ϲ�������" & Chr(13)
    Selection.InsertDateTime DateTimeFormat:=Chr(9) & "EEEE��O��A��", InsertAsField:=False
End Sub

Sub A01_�����ͷ��λ()

'    Application.Run MacroName:="Normal.NewMacros.Macro5"
'    Application.Run MacroName:="A01_�������"
    DW = "��λ"
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

Sub A01_���������ͷ��λ()

End Sub

'    Application.Run MacroName:="Normal.NewMacros.Macro5"
'    Application.Run MacroName:="A01_�������"
    TN = ActiveDocument.Tables.Count
    If TN > 0 Then
    For j = 1 To TN
        ActiveDocument.Tables(j).Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Application.Run MacroName:="A01_�����ͷ��λ"
    Next j
    End If
Sub A01_����ת��Ϊ�ո�()


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

Sub A01_Ӣ������()

'Application.Run MacroName:="����˳��ߵ�"

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

Sub A01_Ӣ������N()

'Application.Run MacroName:="A01_Ӣ������"

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
    Application.Run MacroName:="ҳ������"
    
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

Sub A01_���ת��N()

'Application.Run MacroName:="A01_Ӣ������N"
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


Sub A01_�������()
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
' ���� 2009-4-1 �� ��� ¼��
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
    Selection.TypeText Text:="����ͳ�ƾ�ͳ�����Ϲ�������"
    Selection.TypeParagraph
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdLine, Count:=1
'    ActiveDocument.Shapes.AddTextEffect(msoTextEffect8, "�����", "����", 36#, msoFalse, msoFalse, 244.1, 247.1).Select
    Selection.TypeParagraph
End Sub


Sub �ر���ʽ�Զ�����()
'---------------------------------------------------
'���ܣ��رջ�ĵ�������ʽ���Զ�����
'���ߣ������
'���ڣ�2010��4��28��
'---------------------------------------------------
   
    Dim update As Style
    Set Updates = ActiveDocument.Styles
    For Each update In Updates
        If update.Type = wdStyleTypeParagraph Then
            update.AutomaticallyUpdate = False
        End If
    Next
End Sub

Sub ToggleInterpunction() '��Ӣ�ı�㻥��
Dim ChineseInterpunction() As Variant, EnglishInterpunction() As Variant
Dim myArray1() As Variant, myArray2() As Variant, strFind As String, strRep As String
Dim msgResult As VbMsgBoxResult, N As Byte
'����һ�����ı����������
ChineseInterpunction = Array("��", "��", "��", "��", "��", "��", "��", "����", "-", "��", "��", "��", "��", "��")
'����һ��Ӣ�ı����������
EnglishInterpunction = Array(",", ".", ",", ";", ":", "?", "!", "��", "-", "~", "(", ")", "&lt;", "&gt;")
'��ʾ�û�������MSGBOX�Ի���
msgResult = MsgBox("������Ӣ��㻥����?��Y�����ı��תΪӢ�ı��,��N��Ӣ�ı��תΪ���ı��!", vbYesNoCancel)
Select Case msgResult
Case vbCancel
Exit Sub '����û�ѡ����ȡ����ť,���˳���������
Case vbYes '����û�ѡ����YES,�����ı��ת��ΪӢ�ı��
myArray1 = ChineseInterpunction
myArray2 = EnglishInterpunction
'strFind = " " ( * ) " "
strRep = """\1"""
Case vbNo '����û�ѡ����NO,��Ӣ�ı��ת��Ϊ���ı��
myArray1 = EnglishInterpunction
myArray2 = ChineseInterpunction
strFind = """(*)"""
'strRep = ""\1""
End Select
Application.ScreenUpdating = False '�ر���Ļ����
For N = 0 To UBound(ChineseInterpunction) '��������±굽�ϱ����һ��ѭ��
With ActiveDocument.Content.Find
.ClearFormatting '���޶����Ҹ�ʽ
.MatchWildcards = False '��ʹ��ͨ���
'������Ӧ��Ӣ�ı��,�滻Ϊ��Ӧ�����ı��
.Execute FindText:=myArray1(N), replacewith:=myArray2(N), Replace:=wdReplaceAll
End With
Next
With ActiveDocument.Content.Find
.ClearFormatting '���޶����Ҹ�ʽ
.MatchWildcards = True 'ʹ��ͨ���
.Execute FindText:=strFind, replacewith:=strRep, Replace:=wdReplaceAll
End With
Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub ���ת��E_C()
    Selection.WholeStory
    Dim A As Variant, B As Variant
    
    x2 = Chr(-24157) '��
    x3 = Chr(-23636) '��
    x4 = Chr(-23621) '��
    x5 = Chr(-23622) '��
    x6 = Chr(-23617) '��
    x7 = Chr(-23647) '��
    x8 = Chr(-24150) '"��"
    x9 = Chr(-24149) '"��"
    x10 = Chr(-23640) '"��"
    x11 = Chr(-23639) '"��"
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
Sub ���ת��C_E()
    Selection.WholeStory
    Dim A As Variant, B As Variant
    
    x1 = Chr(-24158) '��
    x2 = Chr(-24157) '��
    x3 = Chr(-23636) '��
    x4 = Chr(-23621) '��
    x5 = Chr(-23622) '��
    x6 = Chr(-23617) '��
    x7 = Chr(-23647) '��
    x8 = Chr(-24150) '"��"
    x9 = Chr(-24149) '"��"
    x10 = Chr(-23640) '"��"
    x11 = Chr(-23639) '"��"
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

Sub A01_ɾ��ĸ()

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

Sub ɨ���ı���׼��()
    
    Dim MyRange As Range
        If Len(Selection.Range.Text) = 0 Then
            Selection.WholeStory
        End If
    Set MyRange = Selection.Range
    
    With MyRange
    
    �س��滻
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Chr(32) & Chr(32) & Chr(32) & Chr(32)
        .Replacement.Text = "^p"
        .Wrap = wdFindStop
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ɾ���ո�K
    ��ǰ�ӿ�
    
    End With

End Sub
Sub A01_��������()
    For Each par In ActiveDocument.Content.Paragraphs
      T = Trim(par.Range.Text)
            If Len(T) = 2 Then
                t1 = Left(par, 1)
                par.Range.Text = t1
            End If
    Next
End Sub

Sub A01_������ȡ����()
    
    Application.DisplayAlerts = wdAlertsNone

Set fs = Application.FileSearch
With fs
    .LookIn = "D:\������"
    .FileName = "*.doc"
    If .Execute(SortBy:=msoSortByFileName, _
    SortOrder:=msoSortOrderAscending) > 0 Then
        For i = 1 To .FoundFiles.Count
    Documents.Open FileName:=.FoundFiles(i)
    On Error Resume Next
    TXT = ActiveDocument.Paragraphs(1).Range.Text
    If Left(TXT, Len(TXT) - 1) = "�й�ͳ�ƿ�����" Then
    
        A01_��ȡ����
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
        MsgBox "û�ҵ������ĵ�"
    End If
End With
    
End Sub

Sub A01_��ȡ����()
    

    A00_��ҳ��ʽ
    Selection.WholeStory
    ɾ���ո�K
    
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
    
    TT1 = "ʡ(��������ֱϽ��)"
    TT2 = "��(�����ݡ���������)"
    TT3 = "��(������������)"
    
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
        If InStr(1, doc.name, "�������.doc", 1) Then
            doc.Activate
            docFound = True
            Exit For
        Else
            docFound = False
        End If
    Next doc
    If docFound = False Then Documents.Open FileName:="D:\00 F2011\�������.doc"
    
    Selection.WholeStory
    Selection.EndKey
    Selection.TypeParagraph
    Selection.TypeText Text:=TT
    Set MyDoc2 = ActiveDocument
    MyDoc1.Activate
    
    

End Sub
Sub ҳ�߾�2����()
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
Sub ��������()
'
' �������� Macro
' ���� 2011-10-17 �� �����: ¼��
'
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.TypeParagraph
    Selection.TypeText Text:="��������"
    Selection.TypeParagraph
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
        "DATE  \@ ""EEEE��O��A��"" ", PreserveFormatting:=True
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



Sub ������������()
    Selection.ParagraphFormat.RightIndent = CentimetersToPoints(0.5)
End Sub
Sub ������1����()
    Selection.ParagraphFormat.RightIndent = CentimetersToPoints(1)
End Sub
Sub ������1���װ�()
    Selection.ParagraphFormat.RightIndent = CentimetersToPoints(1.5)
End Sub
Sub ������2����()
    Selection.ParagraphFormat.RightIndent = CentimetersToPoints(2)
End Sub
Sub Macro8()
'
' Macro8 Macro
' ���� 2013-2-22 �� �����:�ַ� ¼��
'
    CommandBars.Add(name:="�ҵı�񹤾�").Visible = True
End Sub
Sub Macro13()
'
' Macro13 Macro
' ���� 2013-2-22 �� �����:�ַ� ¼��
'
    CommandBars.Add(name:="�ҵı༭����").Visible = True
    CommandBars("�ҵı༭����").Controls.Add Type:=msoControlButton, Before:=1
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
    
    Application.ScreenUpdating = False '�ر���Ļ����
    
    For Each aHyperlink In ActiveDocument.Hyperlinks
        PS(i) = aHyperlink.Address
        WZ(i) = aHyperlink.TextToDisplay
        i = i + 1
    '    MsgBox "i " & i & " ���� " & aHyperlink.Address
        ReDim Preserve PS(UBound(PS) + 1)
        ReDim Preserve WZ(UBound(WZ) + 1)
    Next aHyperlink
        'MsgBox UBound(PS)
    For j = 0 To UBound(PS) - 2
        With MyRange
            .InsertAfter "����" & vbTab & PS(j) & vbTab & WZ(j) & Chr(13)
            '.InsertParagraphAfter
        End With
    Next j
    
    Application.ScreenUpdating = True '����Ļ����
    
End Sub

Sub ��ҳҳ��()
'
' ��ҳҳ�� Macro
' ���� 2013-4-17 �� �����: ¼��
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

Sub B01_ѡ�����ֱ�Ϊ�ϱ�()
    Selection.Font.Superscript = True
    Selection.Font.Color = wdColorBlue
End Sub

Sub B01_���������ı���Ϊ�ϱ�()
    
    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next '���Դ���
    
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
    
    Application.ScreenUpdating = True '����Ļ����

End Sub

Sub B01_��������ת�䴿���Ŀո�()

    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next '���Դ���

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
            KG = KG & "��"
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
    
    A01_ȡ����������
    A01_ȡ����������
    
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Application.ScreenUpdating = True '����Ļ����


End Sub

Sub F01_�������Ÿ帽ע()
'
' ��������������NM -- �ļ����� N -- ͼƬ��  T--�����
    Dim NM As String
        
    ' ȡ�õ�ǰ�ļ���
    Set MyDOC = Application.ActiveWindow.Document
    NM = Left(MyDOC, Len(MyDOC) - 4)
    
    Application.ScreenUpdating = False '�ر���Ļ����

    Documents.Open FileName:="D:\00 Word_Dot\��ע.doc"
    
    Set MyDoc1 = Application.ActiveWindow.Document

    Application.ScreenUpdating = False '�ر���Ļ����

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
    
Application.ScreenUpdating = True '�ر���Ļ����

End Sub
Sub A01_�ӳ�����()

    Dim ML, FN As String
    Set mytable = ActiveDocument.Tables(1)
    R = mytable.Rows.Count
    C = mytable.Columns.Count
    
    Application.ScreenUpdating = False '�ر���Ļ����
    
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

    Application.ScreenUpdating = True '����Ļ����


End Sub


Sub Ӣ������ת��Ϊ��������()
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

Sub A01_ǧ��λ()
'������ּ�ڽ��WORD������ת��Ϊǧ��λ
'�����޶�Ҫ��:-922,337,203,685,477.5808 �� 922,337,203,685,477.5807
'ת�����1000����������ǧ��λ����,С�����Ҳౣ����λС��;1000�������ݲ���
Dim MyRange As Range, i As Byte, myValue As Currency
On Error Resume Next '���Դ���
st = VBA.Timer '��ʱ��
Application.ScreenUpdating = False '�ر���Ļ����

NextFind: Set MyRange = ActiveDocument.Content '����Ϊ���ĵ����ֲ���
With MyRange.Find '����
    .ClearFormatting '�����ʽ
    .Text = "[0-9]{4,15}" '4��15λ����
    .MatchWildcards = True 'ʹ��ͨ���
Do While .Execute 'ÿ�β��ҳɹ�
    i = 2 '��ʼֵΪ2
    '�������С����
    If MyRange.Next(wdCharacter, 1) = "." Then
    '����һ��δ֪ѭ��
        While MyRange.Next(wdCharacter, i) Like "#"
            i = i + 1 'ֻҪ��[0-9]�����������ۼ�
        Wend
        '���¶���RANGE����
        MyRange.SetRange MyRange.start, MyRange.End + i - 1
    End If
    myValue = VBA.Val(MyRange) '�������ת��Ϊ����,Ҳ��ʡ��
    MyRange = VBA.Format(myValue, "Standard") 'תΪǧ��λ��ʽ
    GoTo NextFind 'ת��ָ����
    Loop
    End With

Application.ScreenUpdating = True '�ָ���Ļ����
MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��" '��ʾ�������������ѵ�ʱ��

End Sub

Sub A01_����ͼƬ��С����() '����ͼƬ��СΪ��ǰ�İٷֱ�
Dim N 'ͼƬ����
Dim picwidth
Dim picheight
If Selection.Type = wdSelectionNormal Then
On Error Resume Next '���Դ���
For N = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes����ͼƬ
picheight = ActiveDocument.InlineShapes(N).Height
picwidth = ActiveDocument.InlineShapes(N).Width
ActiveDocument.InlineShapes(N).Height = picheight * 0.5 '���ø߶�
ActiveDocument.InlineShapes(N).Width = picwidth * 0.5 '���ÿ��
Next N
For N = 1 To ActiveDocument.Shapes.Count 'Shapes����ͼƬ
picheight = ActiveDocument.Shapes(N).Height
picwidth = ActiveDocument.Shapes(N).Width
ActiveDocument.Shapes(N).Height = picheight * 0.5 '���ø߶ȱ���
ActiveDocument.Shapes(N).Width = picwidth * 0.5 '���ÿ�ȱ���
Next N

Else: End If
End Sub


Sub A01_����ͼƬ��Сֵ() '����ͼƬ��СΪ�̶�ֵ
Dim N 'ͼƬ����
On Error Resume Next '���Դ���
For N = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes����ͼƬ
ActiveDocument.InlineShapes(N).Height = 200 '����ͼƬ�߶�Ϊ 400px
ActiveDocument.InlineShapes(N).Width = 300 '����ͼƬ��� 300px
Next N
For N = 1 To ActiveDocument.Shapes.Count 'Shapes����ͼƬ
ActiveDocument.Shapes(N).Height = 200 '����ͼƬ�߶�Ϊ 400px
ActiveDocument.Shapes(N).Width = 300 '����ͼƬ��� 300px
Next N
End Sub

Sub A01_ͼƬ��ʽת��()
'* ����������������������������������������������������������
'* Created By SHOUROU@ExcelHome 2007-12-11 5:28:26
'��������System: Windows NT Word: 11.0 Language: 2052
'�� 0281^The Code CopyIn [ThisDocument-ThisDocument]^'
'* ����������������������������������������������������������
'Option Explicit Dim oShape As Variant, shapeType As WdWrapType
On Error Resume Next
If MsgBox("Y��ͼƬ��Ƕ��ʽתΪ����ʽ,N��ͼƬ�ɸ���ʽתΪǶ��ʽ", 68) = 6 Then
shapeType = Val(InputBox(Prompt:="������ͼƬ��ʽ:0=������,1=������, " & vbLf & _
"3=���������·�,4=���������Ϸ�", Default:=0))
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
.WrapFormat.AllowOverlap = False '�������ص�
End With
Next
Else
For Each oShape In ActiveDocument.Shapes
oShape.ConvertToInlineShape
Next
End If
End Sub

Sub A01_GetChineseNum2()
'������ת��Ϊ���ִ�д�����
Dim Numeric As Currency, IntPart As Long, DecimalPart As Byte, MyField As Field, Label As String
Dim Jiao As Byte, Fen As Byte, Oddment As String, Odd As String, MyChinese As String
Dim strNumber As String
Const ZWDX As String = "Ҽ��������½��ƾ���" '����һ�����Ĵ�д���ֳ���
On Error Resume Next '�������
If Selection.Type = wdSelectionNormal Then

With Selection
strNumber = VBA.Replace(.Text, " ", "")
Numeric = VBA.Round(VBA.CCur(strNumber), 2) '�������뱣��С�������λ
'�ж��Ƿ��ڱ����
If .Information(wdWithInTable) Then _
.MoveRight Unit:=wdCell Else .MoveRight Unit:=wdCharacter
'�����ݽ����ж�,�Ƿ���ָ���ķ�Χ��
If VBA.Abs(Numeric) > 2147483647 Then MsgBox "��ֵ������Χ!", _
vbOKOnly + vbExclamation, "Warning": Exit Sub
IntPart = Int(VBA.Abs(Numeric)) '����һ��������
Odd = VBA.IIf(IntPart = 0, "", "Բ") '����һ��STRING����
'�������Ĵ�дǰ�ı�ǩ
Label = VBA.IIf(Numeric = VBA.Abs(Numeric), "����ҽ���д��", "����ҽ���д����")
'��С��������λ��������
DecimalPart = (VBA.Abs(Numeric) - IntPart) * 100
Select Case DecimalPart
Case Is = 0 '�����0,����ѡ��������Ϊ����
Oddment = VBA.IIf(Odd = "", "", Odd & "��")
Case Is < 10 '<10,������ͷ�Ƿ�
Oddment = VBA.IIf(Odd <> "", "Բ��" & VBA.Mid(ZWDX, DecimalPart, 1) & "��", _
VBA.Mid(ZWDX, DecimalPart, 1) & "��")
Case 10, 20, 30, 40, 50, 60, 70, 80, 90 '����ǽ���
Oddment = "Բ" & VBA.Mid(ZWDX, DecimalPart / 10, 1) & "����"
Case Else '���н�,���зֵ����
Jiao = VBA.Left(CStr(DecimalPart), 1) 'ȡ�ý���ֵ
Fen = VBA.Right(CStr(DecimalPart), 1) 'ȡ�÷���ֵ
Oddment = Odd & VBA.Mid(ZWDX, Jiao, 1) & "��" 'ת��Ϊ�ǵ����Ĵ�д
Oddment = Oddment & VBA.Mid(ZWDX, Fen, 1) & "��" 'ת��Ϊ�ֵ����Ĵ�д
End Select
'ָ������������Ĵ�д��ʽ����
Set MyField = .Fields.Add(Range:=.Range, Text:="= " & IntPart & " \*CHINESENUM2")
MyField.Select 'ѡ����(�������ָ���ı�����ѡ������)
'������нǷ������,MychineseΪ""
MyChinese = VBA.IIf(MyField.Result <> "��", MyField.Result, "")
.Text = Label & MyChinese & Oddment
End With
Else
MsgBox "��û��ѡ�����֣���ѡ�����֣�"
End If
End Sub

Sub A01_��Ӣ�ı�㻥��() '��Ӣ�ı�㻥��
Dim ChineseInterpunction() As Variant, EnglishInterpunction() As Variant
Dim myArray1() As Variant, myArray2() As Variant, strFind As String, strRep As String
Dim msgResult As VbMsgBoxResult, N As Byte
'����һ�����ı����������
ChineseInterpunction = Array("��", "��", "��", "��", "��", "��", "��", "����", "��", "��", "��", "��", "��", "��")
'����һ��Ӣ�ı����������
EnglishInterpunction = Array(",", ".", ",", ";", ":", "?", "!", "��", "-", "~", "(", ")", "&lt;", "&gt;")
'��ʾ�û�������MSGBOX�Ի���
msgResult = MsgBox("������Ӣ��㻥����?��Y�����ı��תΪӢ�ı��,��N��Ӣ�ı��תΪ���ı��!", vbYesNoCancel)
Select Case msgResult
Case vbCancel
Exit Sub '����û�ѡ����ȡ����ť,���˳���������
Case vbYes '����û�ѡ����YES,�����ı��ת��ΪӢ�ı��
myArray1 = ChineseInterpunction
myArray2 = EnglishInterpunction
strFind = "��(*)��"
strRep = """\1"""
Case vbNo '����û�ѡ����NO,��Ӣ�ı��ת��Ϊ���ı��
myArray1 = EnglishInterpunction
myArray2 = ChineseInterpunction
strFind = """(*)"""
strRep = "��\1��"
End Select
Application.ScreenUpdating = False '�ر���Ļ����
For N = 0 To UBound(ChineseInterpunction) '��������±굽�ϱ����һ��ѭ��
With ActiveDocument.Content.Find
.ClearFormatting '���޶����Ҹ�ʽ
.MatchWildcards = False '��ʹ��ͨ���
'������Ӧ��Ӣ�ı��,�滻Ϊ��Ӧ�����ı��
.Execute FindText:=myArray1(N), replacewith:=myArray2(N), Replace:=wdReplaceAll
End With
Next
With ActiveDocument.Content.Find
.ClearFormatting '���޶����Ҹ�ʽ
.MatchWildcards = True 'ʹ��ͨ���
.Execute FindText:=strFind, replacewith:=strRep, Replace:=wdReplaceAll
End With
Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub ����ͼƬ��СΪԭʼ��С()
Dim N 'ͼƬ����
Dim picwidth
Dim picheight
On Error Resume Next '���Դ���
For N = 1 To ActiveDocument.InlineShapes.Count 'InlineShapes����ͼƬ
ActiveDocument.InlineShapes(N).Reset
Next N
For N = 1 To ActiveDocument.Shapes.Count 'Shapes����ͼƬ
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
.Title = "��ѡ����һ�ļ���ȷ����������ȫ��WORD�ĵ�"
If .Show <> -1 Then Exit Sub
st = Timer
mypath = .InitialFileName
End With

Application.ScreenUpdating = False
If Dir(mypath & "���Ϊ", vbDirectory) = "" Then MkDir mypath & "���Ϊ" '���Ϊ�ĵ��ı���λ��
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
.SaveAs mypath & "���Ϊ\" & docname & ".doc"
N = N + 1
.Close
End With
End If
Next
End If
End With
MsgBox "��������" & fs.FoundFiles.Count & "���ĵ���������Ŀ���ļ��е�����Ϊ�����Ϊ������һ���ļ����С�" _
& vbCrLf & "����ʱ�䣺" & Format(Timer - st, "0") & "�롣"
Application.ScreenUpdating = True
Exit Sub

hd:
MsgBox "���г������⣬������ֹ��" & vbCrLf & "�Ѵ����ĵ�����" & N _
& vbCrLf & "�����ĵ���" & vbCrLf & fs.FoundFiles(i)
If Not MyDOC Is Nothing Then MyDOC.Close
End Sub


Sub A01_��ʽ()
'
' ��ʽ Macro
' ����ѡ������,��ݼ�Ϊ"Alt+F"
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
MsgBox "��û��ѡ�����֡�"
End If
'
End Sub

Sub A01_��()
'
' �� Macro
' ����ѡ����������ĸ�ϼӻ�
Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdExtend
If Selection.Type = wdSelectionNormal Then
Selection.Font.Italic = True
Selection.Cut
Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
PreserveFormatting:=False
Selection.Delete Unit:=wdCharacter, Count:=1
Selection.TypeText Text:="eq \o(\s\up5(��"
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
MsgBox "��û��ѡ�����֡�"
End If
'
End Sub

Sub A01_Password()
'
' �ļ��Զ�������롣
'
If ActiveDocument.WriteReserved = False Then
If MsgBox("�Ƿ�Ϊ���ĵ�������룿", vbYesNo) = vbYes Then With ActiveDocument
.Password = "123456"
.WritePassword = "123456"
End With

Else
End If
Else
End If
End Sub

Sub A01_Example()

    '�����ĵ��ַ������ظ�Ƶ�������ַ�������
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
    MsgBox "����" & iCount & "�����ظ����ַ�,��ʱ" & VBA.Format(Timer - st, "0.00") & "��"

End Sub

Sub A01_Test()

Dim bw, sw, i As Integer
Dim MyCell As Cell
Dim TXT, MyTXT As Variant
    st = VBA.Timer
    
    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next '���Դ���
    
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
    
    Application.ScreenUpdating = True '�ָ���Ļ����
    
    MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��"


End Sub

Sub A01_Test1()

Dim bw, sw, i As Integer
Dim MyCell As Cell
Dim TXT, MyTXT As Variant
    st = VBA.Timer
    
    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next '���Դ���
    
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
    
    Application.ScreenUpdating = True '�ָ���Ļ����
    
    MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��"

End Sub
Sub A01_Test2()

Dim bw, sw, i As Integer
Dim MyCell As Cell
Dim TXT, MyTXT As Variant
    st = VBA.Timer
    
    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next '���Դ���
    
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
    
    Application.ScreenUpdating = True '�ָ���Ļ����
    
    MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��"

End Sub

Sub A01_��ַ����()  '���ܣ�����ַ���֡�·���ŷ���
    
    Application.ScreenUpdating = False '�ָ���Ļ����
    On Error Resume Next

    Dim MyRange As Range
        If Len(Selection.Range.Text) = 0 Then Selection.WholeStory
    Set MyRange = Selection.Range
    street = "��"
    road = "·"
    txt2 = "��"
    TT = ""
    st = VBA.Timer
    
    With MyRange
        pn = MyRange.Paragraphs.Count
        For i = 1 To pn
            T = Trim(MyRange.Paragraphs(i).Range.Text)
            T = Left(T, Len(T) - 1)
            '���ҽֻ�·
            p2 = InStr(1, T, street, 1)
            If p2 = 0 Then p2 = InStr(1, T, road, 1)
            
            FK1 = Left(T, p2)
            'FK2 = Right(t, Len(t) - p2)
            FK2 = Mid(T, p2 + 1)
            '��"��"
            p3 = InStr(1, FK2, txt2, 1)
            FK3 = Left(FK2, p3)
            FK4 = Mid(FK2, p3 + 1)
            'FK4 = Right(FK2, Len(FK2) - p3)
            TT = TT & FK1 & Chr(9) & FK3 & Chr(9) & FK4 & Chr(13)
        Next i
    End With
    Application.ScreenUpdating = True '�ָ���Ļ����
    ActiveDocument.Content.Text = TT
    MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��"

End Sub

Sub A01_�����޸�һ��ָ���ļ����µ������ļ�()  '���ܣ����������ϰ�ͳһҪ����б�׼��������4�ڡ�����5��

    '���ó������л���
    Application.ScreenUpdating = False '�ر���Ļ���£���߳��������ٶ�
    On Error Resume Next '�����������֮
    st = VBA.Timer '������
    
    '��һ����������׼���õ���������Ŀ¼�µ�ȫ����Ŀ¼��Ϣ��ȡ�����������
    Dim ML() As Variant '������ΪML���������
    Dim i, j As Integer '����2����������
    Documents.Open FileName:="F:\AB\Ŀ¼��Ϣ2.doc"   '���ļ�Ŀ¼��Ϣ.doc
    i = 1 '��ʼ����������i
    For Each PA In ActiveDocument.Paragraphs '����һ��ѭ��������������ǰ�ĵ���ÿ������
        TT = Left(Trim(PA.Range.Text), Len(Trim(PA.Range.Text)) - 1) '��ȡÿ���ı������������з�
        ReDim Preserve ML(i) '���¶����������ά�ȣ�����ԭ��������
        ML(i - 1) = TT  '�Ѷ�ȡ���ı���Ϣ��ֵ������Ԫ��
        i = i + 1  '��������1
    Next PA '��ת����һ�����䣬�����ظ�����
    Documents.Close SaveChanges:=wdDoNotSaveChanges '�ر��ĵ�
    
    '�ڶ�������ÿ����Ŀ¼�µ��ļ����в�������һ���ļ���������ʽ
    For j = 0 To UBound(ML) 'ѭ�����������������е�ÿ��Ԫ��
        
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
                        A01_��ʽ����
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
    Application.ScreenUpdating = True '�ָ���Ļ����
    Application.Visible = True '�ָ��ĵ�����
    ChangeFileOpenDirectory "D:\00 F2013"
    
    MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��" '��ʾ�������������ѵ�ʱ��

End Sub


Sub A01_��ʽ����()

    Selection.WholeStory 'ѡ��ȫ������
        Selection.Font.name = "����"
        Selection.Font.Size = 9
        Selection.Font.name = "Times New Roman"
        Selection.Font.Color = wdColorBlack
        Selection.WholeStory 'ѡ��ȫ������
        Selection.HomeKey Unit:=wdStory
    Selection.Paragraphs(1).Range.Select 'ѡ�б�����
        Selection.Font.name = "����"
        Selection.Font.Size = 14
        Selection.Font.Bold = True
        Selection.Font.Color = wdColorBlack

End Sub

Sub A01_�޸�ָ��Ŀ¼���ļ�()  '���ܣ����Ժ�����

    'A01_�����޸�һ��ָ���ļ����µ������ļ�
        '���ó������л���
    Application.ScreenUpdating = False '�ر���Ļ���£���߳��������ٶ�
    On Error Resume Next '�����������֮
    st = VBA.Timer '������

            Set fs = Application.FileSearch
            With fs
                .LookIn = "F:\��������\���������\61\"
                .FileName = "*.DOC"
                If .Execute(SortBy:=msoSortByFileName, SortOrder:=msoSortOrderAscending) > 0 Then
                    For i = 1 To .FoundFiles.Count
                        FN = .FoundFiles(i)
                        Documents.Open (FN), Visible:=False
                        Documents(FN).Activate
                            A01_��ʽ����
                            Documents(FN).Close SaveChanges:=wdSaveChanges

                    Next i
                End If
            End With
    Application.ScreenUpdating = True '�ָ���Ļ����
    Application.Visible = True '�ָ��ĵ�����
    
    MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��" '��ʾ�������������ѵ�ʱ��

End Sub
Sub A01_����02()  '���ܣ����Ժ�����

    Application.ScreenUpdating = False '�ر���Ļ���£���߳��������ٶ�
    On Error Resume Next '�����������֮
    st = VBA.Timer '������
            
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
                        A01_��ʽ����
                        ActiveDocument.Save
                        Documents.Close SaveChanges:=wdSaveChanges
                    Next i
                End If
            End With
    Application.ScreenUpdating = True '�ָ���Ļ����
    Application.Visible = True '�ָ��ĵ�����
    ChangeFileOpenDirectory "D:\00 F2013\"  '�ָ����ļ���ȱʡ·��
    
    MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��" '��ʾ�������������ѵ�ʱ��

End Sub

Sub TT_Ontime()

'��ʾ������ 15 ������� my_Procedure ���̣������ڿ�ʼ��ʱ��

    Application.OnTime Now + TimeValue("00:00:15"), "my_Procedure"

End Sub

Sub A01_�൥��()

   '�����д��Frank
   '��д���ڣ�2014��3��20��
   '�����ܣ���ָ��Ҫ����ض����11����ʱ����б༭��ָ��Ҫ��
   '1.ɾ�����ĵ�7��8��9��10�У�
   '2.�ڱ�����������1�հ��У�
   '3.��������ﵥԪ���У��У������ȫ�����ݶ���0����ȥ��
   '4.����ڶ��䣬��ǰ�κ��о�Ϊ�̶�ֵ10��
   
   Application.ScreenUpdating = False '�ر���Ļ����, ����ߴ��������ٶ�
   'st = VBA.Timer '��ʱ��
   Dim MyTab As Table '����һ����ΪMyTab�ı�����
   
   '��һ�����жϲ���㣨�����λ�ã��Ƿ��ڱ���ڣ�
   '����ǣ���ִ�г�����룻
   '������򵯳���ʾ���ڣ���ʾ�û�����������ڱ���ڣ����ⵥԪ���ڣ�
   
   If Selection.Information(wdWithInTable) = True Then
        Set MyTab = Selection.Tables(1)
        CN = MyTab.Columns.Count 'ʶ���������
        RN = MyTab.Rows.Count 'ʶ���������
        If CN <> 11 Then '�������11�еı���򵯳���ʾ���ڣ���ʾ�û����Ա�����κδ����˳�������
            MsgBox "����񲻷���ָ��������11����ʱ�������ִ�б��������ѡ����������ı��лл��"
        Else
            MyTab.Rows(RN).Select 'ѡ�����һ��
            Selection.InsertRowsBelow 1 '����1��
            MyTab.Columns(7).Select
            Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
            Selection.Cut 'ɾ�����ĵ�7��8��9��10��
            MyTab.Select
            With Selection.ParagraphFormat '����ڶ��䣬��ǰ�κ��Ϊ0���о�Ϊ�̶�ֵ10��
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
            '���ñ����Ϊ�������Զ�����
            MyTab.AutoFitBehavior (wdAutoFitContent)
            MyTab.AutoFitBehavior (wdAutoFitContent)
            '�ѹ�궨λ�����һ�еĵ�2����Ԫ����
            MyTab.Rows(RN + 1).Select 'ѡ�����һ��
            Selection.MoveLeft Unit:=wdCharacter, Count:=1
            Selection.MoveRight Unit:=wdCharacter, Count:=1
        End If
   Else
        MsgBox "��ע�⡿����㣨��꣩���ڱ���У�" & Chr(13) & _
           "���������뽫����㣨��꣩���ڱ�������ⵥԪ���У�" & Chr(13) & _
           "��������Ȼ����ִ�б������лл��"
   End If
   
      Application.ScreenUpdating = True '�ָ���Ļ����
      'MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��" '��ʾ�������������ѵ�ʱ��

End Sub

Sub A01_����()

    '�����д��Frank
    '��д���ڣ�2014��3��20��
    '�����ܣ���ָ��Ҫ���һ���ض����11����ʱ����б༭
    
    Application.ScreenUpdating = False '�ر���Ļ����, ����ߴ��������ٶ�
    st = VBA.Timer '��ʱ��
    Dim TN As Integer '����һ����tn�ı���,���ڼ����������
    On Error Resume Next '�����������֮
    
    TN = ActiveDocument.Tables.Count
    
    For i = 1 To TN
        ActiveDocument.Tables(i).Select
        A01_�൥��
    Next i
      Application.ScreenUpdating = True '�ָ���Ļ����
      MsgBox "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��" '��ʾ�������������ѵ�ʱ��
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
    newTextbox.TextFrame.TextRange = "�۷�"
End Sub

Sub A01_���������ͷ()
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
        If Left(A1, 2) = "��λ" Then
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.Expand Unit:=wdParagraph
            If Len(Selection.Range.Text) < 2 Then Selection.Delete
            Selection.MoveUp Unit:=wdLine, Count:=1
            Selection.Expand Unit:=wdParagraph
            Selection.MoveDown Unit:=wdLine, Count:=1, Extend:=wdExtend
            B01_ѡ�����ֱ�Ϊ��ͷ
        Else
            B01_ѡ�����ֱ�Ϊ��ͷ
        End If
        Next i
    End If
End Sub

Sub A01_���������ͷE()
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
        B01_ѡ�����ֱ�Ϊ��ͷE
    Else
        B01_ѡ�����ֱ�Ϊ��ͷE
    End If
    
    Next i
    
    End If
End Sub


Sub A01_�����Ӵֱ��ϼ���()
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
                .Execute FindText:="��"
                If .Found = True Then
                    .Parent.Expand Unit:=wdParagraph
                End If
            End With
            A1 = Selection.Range.Text
            P1 = InStr(1, A1, "��", 1)
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
                .Execute FindText:="��"
                If .Found = True Then
                    .Parent.Expand Unit:=wdParagraph
                End If
            End With
            A1 = Selection.Range.Text
            P1 = InStr(1, A1, "��", 1)
            If P1 > 0 Then
                Selection.SelectRow
                Selection.Range.Font.Bold = True
            End If
        Next i
    End If
    
End Sub

Sub A01_�����Ӵֱ��ϼ���E()
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

Sub A01_�����Ӵֱ���е��ض���()
    '����д��С�һ���������������������������мӴ֣����Ӵ�������ʮ��
    Dim A As Variant
    A = Array("һ", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ" _
          , "ʮһ", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "ʮ��", "��ʮ")
    Selection.Tables(1).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.SelectColumn
    Set MyRange = Selection.Range
    
    For j = 0 To UBound(A)
        With Selection.Find
            .ClearFormatting
            .Execute FindText:=A(j) & "��"
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
            .Execute FindText:=A(j) & "��"
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

Sub A01_�����Ӵֱ���е��ض���E()
    '����д��С�һ���������������������������мӴ֣����Ӵ�������ʮ��
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


Sub A01_ѡ�������ʽ�̶��о�12��()
    
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 12
    End With
    
End Sub

Sub A01_����Ŀ¼()
    
    Dim P1 As String
    Dim TN As Integer
    
    P1 = ActiveDocument.Paragraphs(1).Range.Text
    P1 = Left(P1, Len(P1) - 1)
    TN = ActiveDocument.Tables.Count
    
    If TN = 3 And P1 = "��������ͳ�Ƶ�����ĿĿ¼" Then
    
    ��ҳҳ��
    
    ActiveDocument.Tables(1).Select
    A01_������ĿĿ¼����ʽ����
    ActiveDocument.Tables(2).Select
    A01_������ĿĿ¼����ʽ����
    ActiveDocument.Tables(3).Select
    A01_������ĿĿ¼����ʽ����
    
    Else
    
    MsgBox "��Ǹ�����ĵ�������ִ�б�������������"
    End If
    
End Sub

Sub A01_������ĿĿ¼����ʽ����()

    ���B
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

Sub A01_����Ŀ¼�ĵ�()
    Dim MyDOC As Document, MyDir As String
    Set MyDOC = Documents.Add
    ��ҳҳ��
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
        .NameFarEast = "����"
        .NameAscii = "����"
        .Size = 12
        .Bold = True
    End With
    MyDOC.Paragraphs(1).Range.Select
    Selection.TypeText Text:="��������ͳ�Ƶ�����ĿĿ¼"
    MyDOC.Paragraphs(1).Range.Select
    Selection.Font.Color = wdColorRed
    
    MyDOC.Paragraphs(3).Range.Select
    Selection.TypeText Text:="���ҷ�չ�͸ĸ�ίԱ��"
    MyDOC.Paragraphs(3).Range.Select
    Selection.Font.Color = wdColorBlue

    MyDOC.Paragraphs(5).Range.Select
    Selection.TypeText Text:="������Ŀһ����"
    MyDOC.Paragraphs(5).Range.Select
    Selection.Font.Color = wdColorBlack
    
    MyDOC.Paragraphs(9).Range.Select
    Selection.TypeText Text:="������Ŀһ����"
    MyDOC.Paragraphs(9).Range.Select
    Selection.Font.Color = wdColorBlack
    


    'Selection.Sections(1).Footers(1).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberOutside, FirstPage:=True
   ' MyDoc.SaveAs FileName:=FN, FileFormat:=wdFormatDocument    '�����ļ�

End Sub

Sub A01_�����滻()
'
' ȫ������ת��S Macro
' ���� 2003-6-27 �� DHG ¼��
'
    Selection.WholeStory
    Dim A As Variant
    Dim B As Variant
    A = Array("һ", "��", "��", "��", "��", "��", "��")
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

Sub A01_ɾ�ո�س�()

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

Sub A01_���ϲ���ͷ��ע()

    N = ActiveDocument.Tables.Count
    If N > 0 Then
    For i = 1 To N
        ActiveDocument.Tables(i).Select
        Selection.MoveLeft Unit:=wdCharacter, Count:=1
        Selection.SelectRow
            If Left(Selection.Range.Text, 1) = "��" Then
                Selection.Cut
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.PasteSpecial DataType:=wdPasteText
            End If
        ActiveDocument.Tables(i).Select
        Selection.EndKey
        Selection.SelectRow
        If InStr(1, Selection.Range.Text, "ע��", 1) > 0 Then
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
            If InStr(1, Selection.Range.Text, "��λ��", 1) > 0 Then
                Selection.Cut
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
                Selection.TypeParagraph
                Selection.PasteSpecial DataType:=wdPasteText
            End If
    Next i
    End If

End Sub

Sub A01_���ϲ���ͷ��עE()

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

Sub A01_���ÿ�ݼ�()

    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyV, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="ճ���ı�"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyK, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="���Ŀո�"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyZ, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="A00_��ҳ��ʽ"
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyQ, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="��׼��ҳ��ʽW"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyB, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="���B"
        
End Sub

Sub A01_βע���()

    Dim RNG As Range
    Set RNG = ActiveDocument.Range
    TXT = RNG.Text
    TXT = Left(TXT, Len(TXT) - 1)
    L = Len(TXT)
    C = "[ ]"
    i = 1
    P = InStr(1, TXT, C, 1)
    
    Application.ScreenUpdating = False '�ر���Ļ����
    
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
    
    Selection.HomeKey Unit:=wdStory '����������ĵ���ʼ��
    Application.ScreenUpdating = True '�ָ���Ļ����

End Sub
Sub A01_ɾ��ȫ��βע()

    Dim N As Integer
    N = ActiveDocument.Endnotes.Count
    
    If N > 0 Then
        For Each nt In ActiveDocument.Endnotes
            nt.Delete
        Next nt
    End If

End Sub


Sub A01_��ʾβע����()

    Dim N As Integer
    N = ActiveDocument.Endnotes.Count
    MsgBox N
    
    If N > 0 Then
        For Each nt In ActiveDocument.Endnotes
            MsgBox nt.Range.Text
        Next nt
    End If

End Sub

Sub A01_ͼƬ��ʽ�ɸ�����ת��ΪǶ����()

    On Error Resume Next
    Dim N As Integer
    Application.ScreenUpdating = False '�ر���Ļ����
    N = ActiveDocument.Shapes.Count
    Selection.HomeKey Unit:=wdStory
    If N > 0 Then
        For Each ishape In ActiveDocument.Shapes
            ishape.Select
            ishape.ConvertToInlineShape
        Next ishape
    End If
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '�ָ���Ļ����
    MsgBox "��ת����" & N & "��ͼƬΪǶ����"

End Sub

Sub A01_ͼƬ��ʽ��Ƕ����ת��Ϊ������()
    On Error Resume Next
    Dim N As Integer
    Application.ScreenUpdating = False '�ر���Ļ����
    Selection.HomeKey Unit:=wdStory
    N = ActiveDocument.InlineShapes.Count
    If N > 0 Then
        For Each ishape In ActiveDocument.InlineShapes
            ishape.Select
            ishape.ConvertToShape
        Next ishape
    End If
    Selection.HomeKey Unit:=wdStory
    Application.ScreenUpdating = True '�ָ���Ļ����
    MsgBox "��ת����" & N & "��ͼƬΪ����������"
    
End Sub

Sub A01_���Ժ�����2()  '���ܣ����Ժ�����

 A01_ɾ��ǧ��λ����
 
'A01_ͼƬ��ʽ�ɸ�����ת��ΪǶ����
'A01_ͼƬ��ʽת��
    
End Sub


Sub A01_ȡ����������()
    
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
Sub A01_ɾ��ǧ��λ����()

    Dim MyRange As Range
    Dim T As String, S1 As String, S2 As String, S3 As String, S4 As String, TT As String
    Dim P As Integer, N As Integer, i As Integer, j As Integer
    
    Application.ScreenUpdating = False '�ر���Ļ����
    
    '�趨�������÷�ΧΪ�û�ѡ������������û�û��ѡ��������Ĭ��ѡ������Ϊ��ƪ�ĵ�
    If Len(Selection.Range.Text) = 0 Then Selection.WholeStory
    Set MyRange = Selection.Range
    t1 = "��" 'ָ���ַ��������Ƕ���
    TT = ""  '����һ�����ַ���
    st = VBA.Timer '��ʱ��
    j = 0 '��ʼ��������
    N = MyRange.Paragraphs.Count
    
        For i = 1 To N
            T = MyRange.Paragraphs(i).Range.Text
            T = Left(T, Len(T) - 1)  '�����ı����������з����س���
            P = InStr(1, T, t1, 1)   'ָ���ֺ��ڶ����е�λ��
            TT = ""
            
            If P > 0 Then
            Do Until P = 0
                If InStr(1, T, t1, 1) > 0 Then
                    P = InStr(1, T, t1, 1)
                    S1 = Left(T, P - 1)  ' ָ���ַ�ǰ����ı�
                    S2 = Right(T, Len(T) - P)   ' ָ���ַ�������ı�
                    S3 = Right(S1, 1) ' ָ���ַ�ǰ��һ���ַ�
                    S4 = Left(S2, 1)  ' ָ���ַ�����һ���ַ�
                    '���ָ���ַ�ǰ��һ���ַ��ͺ���һ���ַ���Ϊ���֣���ɾ���ö���
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
    Selection.HomeKey Unit:=wdStory '����������ĵ���ʼ��
    Application.ScreenUpdating = True '�ָ���Ļ����
    MsgBox "���滻  " & j & "  ����" & Chr(13) & "����������ʱ" & VBA.Format(Timer - st, "0.00") & "��" '��ʾ�������������ѵ�ʱ��

End Sub

Sub A01_�������Ƿ�����()
    Selection.TypeText Text:=Chr(-24142) & Year(Date) & Chr(-24141)
End Sub

Sub A01_����ĵ�()

    Application.ScreenUpdating = False '�ر���Ļ����

    Dim NM As String   'NM -- �ļ���
    Dim N As Integer   'N -- �ļ���
    Dim TXT As String  'TXT -- �ı�
    Dim FN() As Variant  '��������
    Dim TB As Table  '���������
    Dim RS As Range   '��ʼ��
    Dim RE As Range   '������
    Dim MyRange As Range '��Ҫ����������
    
    ChangeFileOpenDirectory "D:\01 MyFiles"  '���ô��ļ���·��
    Documents.Open ("Ŀ¼.doc")
    Set TB = ActiveDocument.Tables(1)
    TXT = TB.Cell(1, 1).Range.Text
    TXT = Left(TXT, Len(TXT) - 1)
    ReDim Preserve FN(1)  '���¶�������FN
    FN(0) = TXT  '��ʼ������Ԫ��
    For i = 2 To TB.Rows.Count
        TXT = TB.Cell(i, 1).Range.Text
        TXT = Left(TXT, Len(TXT) - 1)
        ReDim Preserve FN(UBound(FN) + 1)
        FN(UBound(FN) - 1) = TXT
    Next i
    
    'MsgBox FN(UBound(FN) - 1)
   ' MsgBox UBound(FN)
    
    Documents.Close SaveChanges:=wdDoNotSaveChanges  '�ر�Ŀ¼�ĵ�
    
    Documents.Open ("�����ĵ�.doc")  '�򿪲����ĵ�
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
    Application.ScreenUpdating = True '����Ļ����
    ChangeFileOpenDirectory "D:\00 F2013\" '�ָ����ļ���ȱʡ·��
    
End Sub


Sub A00_����ͼƬת��ΪǶ��ʽͼƬ()
    
    N = ActiveDocument.InlineShapes.Count ' ȡ���ĵ���ͼƬ��
    MsgBox N
    
    N_SHP = ActiveDocument.Shapes.Count
    
    MsgBox N_SHP
    
    '����ĵ����и���ʽͼƬ������ת��ΪǶ��ʽͼƬ
    If ActiveDocument.Shapes.Count > 0 Then
        For Each oShape In ActiveDocument.Shapes
            oShape.ConvertToInlineShape
        Next
    End If
    
    N = ActiveDocument.InlineShapes.Count ' ȡ���ĵ���ͼƬ��
    MsgBox N
    
End Sub
Sub ���ı�ճ��()
'
' ���ı�ճ�� ��
'
'
    CommandBars("Office Clipboard").Visible = False
    Selection.PasteSpecial Link:=False, DataType:=wdPasteText, Placement:= _
        wdInLine, DisplayAsIcon:=False
    Selection.TypeParagraph
    Selection.TypeParagraph
End Sub


Sub G00_ÿ�յ���()

    Dim MyDOC As Document
    Dim MyRange As Range
    
    Application.ScreenUpdating = False '�ر���Ļ����
    On Error Resume Next
    
    '����һ�����ĵ�
    Set MyDOC = Documents.Add
    
    '�����ĵ�����Ϊ����ҳ��
    G00_����ҳ������
    
    '����15������
    For i = 1 To 13
        Selection.TypeParagraph
    Next i
    
    '���ö����ʽ�������ֺ�
    MyDOC.Range.Select
    Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
    Selection.ParagraphFormat.LineSpacing = 30
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.name = "����_GB2312"
    Selection.Font.Size = 15
    
    '���õ�һ�ж����ʽ�������ֺ�
    Set MyRange = MyDOC.Paragraphs(1).Range
    MyRange.Text = "δ�����"
    MyRange.Font.name = "����"
    MyRange.Font.Size = 16
    MyRange.ParagraphFormat.Alignment = wdAlignParagraphRight
    
     '���õڶ��ж����ʽ�������ֺ�
    Set MyRange = MyDOC.Paragraphs(2).Range
    MyRange.Text = "����ת��"
    MyRange.Font.name = "����"
    MyRange.Font.Size = 16
    MyRange.ParagraphFormat.Alignment = wdAlignParagraphRight
    
    '������ͷ����
    Set MyRange = MyDOC.Paragraphs(3).Range
    MyRange.Text = "ÿ�յ���"
    MyRange.Font.name = "�����п�"
    MyRange.Font.Size = 80
    MyRange.Font.Color = wdColorRed
    MyRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    MyRange.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    MyRange.ParagraphFormat.SpaceBefore = 20
    MyRange.ParagraphFormat.LineUnitBefore = 4
    
    '�����ĺ�
    Set MyRange = MyDOC.Paragraphs(4).Range
    MyRange.Text = "��" & Year(Date) & "��� �ڣ�"
    MyRange.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    MyRange.ParagraphFormat.SpaceBefore = 15
    MyRange.ParagraphFormat.SpaceAfter = 15
    MyRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    '���÷��ĵ�λ�ͷ�������
    Set MyRange = MyDOC.Paragraphs(5).Range
    MyRange.ParagraphFormat.TabStops.ClearAll
    MyRange.Select
    Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(15.2), Alignment:=wdAlignTabRight, Leader:=wdTabLeaderSpaces
    Selection.TypeText Text:="����ͳ�ƾְ칫��" & vbTab & Year(Date) & "��" & Month(Date) & "��" & Day(Date) & "��"
    MyRange.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    MyRange.ParagraphFormat.Alignment = wdAlignParagraphLeft
    MyRange.Select
    'Set MyRange = ActiveDocument.Range(start:=Selection.Range.start, End:=ActiveDocument.Range.End - 1)
    Set MyRange = ActiveDocument.Range(start:=Selection.Characters(1).start, End:=Selection.Characters(8).End)
    MyRange.Select
    Selection.Font.name = "����"
    
    '���ú���
    Set MyRange = MyDOC.Paragraphs(6).Range
    MyRange.Borders(wdBorderTop).LineStyle = Options.DefaultBorderLineStyle
    MyRange.Borders(wdBorderTop).LineWidth = wdLineWidth300pt
    MyRange.Borders(wdBorderTop).Color = wdColorRed
    
     '���ñ�������ʽ�������ֺ�
    Set MyRange = MyDOC.Paragraphs(8).Range
    MyRange.Font.name = "����С����_GBK"
    MyRange.Font.Size = 22
    MyRange.Text = "����"
    MyRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
     '�������Ķ����ʽ�������ֺ�
    Set MyRange = MyDOC.Paragraphs(9).Range
    MyRange.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0.35)
    MyRange.ParagraphFormat.CharacterUnitFirstLineIndent = 2
    MyRange.Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeParagraph
    Application.ScreenUpdating = True '�ָ���Ļ����

End Sub
Sub G00_����ҳ������()
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

Sub �ҵĺ�1()
'
' �ҵĺ�1 ��
'
'
    Selection.TypeText Text:="����ͳ�ƾ�"
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

Sub ָ����Χ���Ӵ־���()
    Dim RNG As Range
    Set RNG = ActiveDocument.Paragraphs(1).Range
    With RNG
        .Bold = True
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Font
            .name = "����"
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

Sub �ж��ĵ��Ƿ��()

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


Sub A00_����ĵ���������()

    Dim doc As Document
    Set doc = ActiveDocument
    If doc.ProtectionType <> wdNoProtection Then doc.Unprotect
    
End Sub

Sub A00_���ĵ���������()

    Dim doc As Document
    Set doc = ActiveDocument
    If doc.ProtectionType = wdNoProtection Then doc.Protect Type:=wdAllowOnlyFormFields
    
End Sub
Sub A00_�����ĵ������޶�()

    Dim doc As Document
    Set doc = ActiveDocument
    doc.AcceptAllRevisions
    
End Sub

Sub ��д()
'
' ��д ��
'
'

End Sub
