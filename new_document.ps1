#Requires -Version 7.0

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -Path $env:WINDIR\assembly\GAC_MSIL\office\*\office.dll -PassThru
Get-ChildItem -Path $env:windir\assembly -Recurse -Filter Microsoft.Office.Interop.Word* -File | ForEach-Object {
    Add-Type -LiteralPath ($_.FullName) -PassThru
}

# Функция для преобразования сантиметров в пункты
function Convert-CmToPoint($cm) {
    return $cm / 2.54 * 72
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Выбор полей для документа'
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = 'CenterScreen'

$radioButton1 = New-Object System.Windows.Forms.RadioButton
$radioButton1.Location = New-Object System.Drawing.Point(10,40)
$radioButton1.Size = New-Object System.Drawing.Size(260,20)
$radioButton1.Text = 'Правое поле 1,5 см'

$radioButton2 = New-Object System.Windows.Forms.radioButton
$radioButton2.Location = New-Object System.Drawing.Point(10,70)
$radioButton2.Size = New-Object System.Drawing.Size(260,20)
$radioButton2.Text = 'Правое поле 1 см'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(110,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

$form.Controls.Add($radioButton1)
$form.Controls.Add($radioButton2)
$form.Controls.Add($okButton)

$form.AcceptButton = $okButton
$result = @{}
# Добавление значений в словарь

$form.Add_Shown({$form.Activate()})
$form.ShowDialog() | Out-Null

if ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $result['radioButton1'] = $radioButton1.Checked
    $result['radioButton2'] = $radioButton2.Checked
}

$word = New-Object -ComObject Word.Application
$word.Visible = $true
$doc = $word.Documents.Add()

# Настройка стиля "Обычный"
$normalStyle = $doc.Styles["Обычный"]
$normalStyle.Font.Name = "Times New Roman Cyr"
$normalStyle.Font.Size = 12
$normalStyle.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphJustify
$normalStyle.ParagraphFormat.LineSpacingRule = [Microsoft.Office.Interop.Word.WdLineSpacing]::wdLineSpace1pt5
$normalStyle.ParagraphFormat.FirstLineIndent = 1.25 * 72 / 2.54 # 1.25 см в пунктах
$normalStyle.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorAutomatic

# Настройка стиля "Заголовок 1"
$heading1Style = $doc.Styles["Заголовок 1"]
$heading1Style.Font.Name = "Times New Roman Cyr"
$heading1Style.Font.Size = 12
$heading1Style.Font.Bold = $true
$heading1Style.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphLeft
$heading1Style.ParagraphFormat.LineSpacingRule = [Microsoft.Office.Interop.Word.WdLineSpacing]::wdLineSpace1pt5
$heading1Style.ParagraphFormat.SpaceAfter = 18
$heading1Style.ParagraphFormat.PageBreakBefore = $true
$heading1Style.ParagraphFormat.FirstLineIndent = 1.25 * 72 / 2.54 # 1.25 см в пунктах
$heading1Style.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorAutomatic

# Настройка стиля "Заголовок 2"
$heading2Style = $doc.Styles["Заголовок 2"]
$heading2Style.Font.Bold = $true
$heading2Style.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphLeft
$heading2Style.ParagraphFormat.LineSpacingRule = [Microsoft.Office.Interop.Word.WdLineSpacing]::wdLineSpace1pt5
$heading2Style.ParagraphFormat.SpaceBefore = 18
$heading2Style.ParagraphFormat.SpaceAfter = 18
$heading2Style.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorAutomatic

# Создание нового стиля "Заголовок 1 Дополнительный"
$heading1ExtraStyle = $doc.Styles.Add("Заголовок 1 Дополнительный", [Microsoft.Office.Interop.Word.WdStyleType]::wdStyleTypeParagraph)
$heading1ExtraStyle.BaseStyle = $heading1Style
$heading1ExtraStyle.Font.AllCaps = $true
$heading1ExtraStyle.Font.Bold = $true
$heading1ExtraStyle.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
$heading1ExtraStyle.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorAutomatic
$heading1ExtraStyle.ParagraphFormat.PageBreakBefore = $true
$heading1ExtraStyle.ParagraphFormat.FirstLineIndent = 0

# Создание стиля "Оглавление"
$tocStyle = $doc.Styles.Add("Оглавление", [Microsoft.Office.Interop.Word.WdStyleType]::wdStyleTypeParagraph)
$tocStyle.Font.AllCaps = $true
$tocStyle.Font.Bold = $true
$tocStyle.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
#$tocStyle.ParagraphFormat.PageBreakBefore = $true
$tocStyle.ParagraphFormat.SpaceAfter = 18
$tocStyle.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorAutomatic
$tocStyle.ParagraphFormat.FirstLineIndent = 0

# Создание стиля "Подпись для картинок"
$captionPicStyle = $doc.Styles.Add("Подпись для картинок", [Microsoft.Office.Interop.Word.WdStyleType]::wdStyleTypeParagraph)
$captionPicStyle.Font.Size = 12
$captionPicStyle.Font.Name = "Times New Roman Cyr"
$captionPicStyle.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
$captionPicStyle.ParagraphFormat.SpaceBefore = 6
$captionPicStyle.ParagraphFormat.SpaceAfter = 18
$captionPicStyle.ParagraphFormat.LineSpacingRule = [Microsoft.Office.Interop.Word.WdLineSpacing]::wdLineSpaceSingle
$captionPicStyle.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorAutomatic
$captionPicStyle.ParagraphFormat.FirstLineIndent = 0

# Создание стиля "Подпись для таблиц"
$captionTableStyle = $doc.Styles.Add("Подпись для таблиц", [Microsoft.Office.Interop.Word.WdStyleType]::wdStyleTypeParagraph)
$captionTableStyle.Font.Size = 12
$captionTableStyle.Font.Name = "Times New Roman Cyr"
$captionTableStyle.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphLeft
$captionTableStyle.ParagraphFormat.SpaceBefore = 6
$captionTableStyle.ParagraphFormat.SpaceAfter = 18
$captionTableStyle.ParagraphFormat.LineSpacingRule = [Microsoft.Office.Interop.Word.WdLineSpacing]::wdLineSpaceSingle
$captionTableStyle.ParagraphFormat.FirstLineIndent = 1.25 * 72 / 2.54 # 1.25 см в пунктах
$captionTableStyle.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorAutomatic

# Создание стиля "Картинка"
$picStyle = $doc.Styles.Add("Картинка", [Microsoft.Office.Interop.Word.WdStyleType]::wdStyleTypeParagraph)
$picStyle.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
$picStyle.ParagraphFormat.SpaceBefore = 0
$picStyle.ParagraphFormat.SpaceAfter = 0
$picStyle.Font.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorAutomatic
$picStyle.ParagraphFormat.FirstLineIndent = 0

# Установка полей документа
$doc.PageSetup.TopMargin = Convert-CmToPoint 2
$doc.PageSetup.BottomMargin = Convert-CmToPoint 2
$doc.PageSetup.LeftMargin = Convert-CmToPoint 3
$doc.PageSetup.RightMargin = Convert-CmToPoint ($result['radioButton1'] ? 1.5 : 1)

$Section = $Doc.Sections.Item(1)
#$FirstPageHeader = $Section.Headers.Item([Microsoft.Office.Interop.Word.WdHeaderFooterIndex]::wdHeaderFooterFirstPage)
# Отключение номера страницы на первой странице
$Section.PageSetup.DifferentFirstPageHeaderFooter = $true
$Selection = $Word.Selection

# Добавление текста на титульную страницу
$Selection.TypeText("Титульная страница")
$Selection.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)

# Добавление слова "Оглавление" и применение к нему стиля "Оглавление"
$Selection.TypeText("Оглавление")
$Selection.Style = $Doc.Styles.Item("Оглавление")

# Вставка настраиваемого оглавления
$Doc.TablesOfContents.Add($Selection.Range)

# Добавление слов "Введение" и "Заключение" с применением стиля "Заголовок 1 Дополнительный"
$Selection.TypeText("`nВведение")
$Selection.Style = $Doc.Styles.Item("Заголовок 1 Дополнительный")
$Selection.TypeText("`nЗаключение")
$Selection.Style = $Doc.Styles.Item("Заголовок 1 Дополнительный")

# Вставка номеров страниц внизу по центру
$Footer = $Section.Footers.Item(1)
$Footer.PageNumbers.Add(1)

# Обновление всех полей в документе
$Doc.Fields.Update()

# Сохранение документа
# Получение текущего пути
$currentPath = Get-Location

# Сохранение документа в текущей папке с именем 'Document.docx'
$doc.SaveAs([ref]"$currentPath\Document.docx")

$word.Quit()
