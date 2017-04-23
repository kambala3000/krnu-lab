Attribute VB_Name = "NewMacros"
Sub TitlePage()
Attribute TitlePage.VB_Description = "Создание титульной сраницы"
Attribute TitlePage.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.TitlePage"
'
' TitlePage Макрос
' Создание титульной сраницы
'
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 14
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="МІНІСТРЕСТВО ОСВІТИ І НАУКИ УКРАЇНИ"
    Selection.TypeParagraph
    Selection.TypeText Text:="КРЕМЕНЧУЦЬКИЙ НАЦІОНАЛЬНИЙ УНІВЕРСИТЕТ"
    Selection.TypeParagraph
    Selection.TypeText Text:="ІМЕНІ МИХАЙЛА ОСТРОГРАДСЬКОГО"
    Selection.ParagraphFormat.SpaceAfter = 18
    Selection.TypeParagraph
    Selection.TypeText Text:="КАФЕДРА ІНФОРМАТИКИ І ВИЩОЇ МАТЕМАТИКИ"
    Selection.ParagraphFormat.SpaceAfter = 30
    Selection.TypeParagraph
    Selection.Font.Size = 20
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="РЕФЕРАТ"
    Selection.ParagraphFormat.SpaceAfter = 18
    Selection.TypeParagraph
    Selection.Font.Size = 14
    Selection.Font.Bold = wdToggle
    Selection.TypeText Text:="з дисципліни «Назва дисципліни»"
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.TypeParagraph
    Selection.TypeText Text:="на тему «Назва теми»"
    Selection.ParagraphFormat.SpaceAfter = 120
    Selection.TypeParagraph
    Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
    Selection.TypeText Text:="Студента 0 курсу ГРУПА 00-0 групи"
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.TypeParagraph
    Selection.TypeText Text:="Спеціальності «Спеціальність»"
    Selection.TypeParagraph
    Selection.TypeText Text:="Прізвище І.Б."
    Selection.TypeParagraph
    Selection.TypeText Text:="Керівник: Прізвище І.Б."
    Selection.ParagraphFormat.SpaceAfter = 270
    Selection.TypeParagraph
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.TypeText Text:="КРЕМЕНЧУК 2017"
End Sub
