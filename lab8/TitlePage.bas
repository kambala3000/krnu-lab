Attribute VB_Name = "NewMacros"
Sub TitlePage()
Attribute TitlePage.VB_Description = "�������� ��������� �������"
Attribute TitlePage.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.TitlePage"
'
' TitlePage ������
' �������� ��������� �������
'
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 14
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="̲Ͳ�������� ��²�� � ����� �������"
    Selection.TypeParagraph
    Selection.TypeText Text:="������������� ��ֲ�������� �Ͳ��������"
    Selection.TypeParagraph
    Selection.TypeText Text:="���Ͳ ������� ���������������"
    Selection.ParagraphFormat.SpaceAfter = 18
    Selection.TypeParagraph
    Selection.TypeText Text:="������� ����������� � ���ί ����������"
    Selection.ParagraphFormat.SpaceAfter = 30
    Selection.TypeParagraph
    Selection.Font.Size = 20
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="�������"
    Selection.ParagraphFormat.SpaceAfter = 18
    Selection.TypeParagraph
    Selection.Font.Size = 14
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="� ��������� ������ ���������"
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.TypeParagraph
    Selection.TypeText Text:="�� ���� ������ ����"
    Selection.ParagraphFormat.SpaceAfter = 120
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.TypeText Text:="�������� 0 ����� ����� 00-0 �����"
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.TypeParagraph
    Selection.TypeText Text:="������������ ��������������"
    Selection.TypeParagraph
    Selection.TypeText Text:="������� �.�."
    Selection.TypeParagraph
    Selection.TypeText Text:="�������: ������� �.�."
    Selection.ParagraphFormat.SpaceAfter = 270
    Selection.TypeParagraph
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="��������� 2017"
End Sub
