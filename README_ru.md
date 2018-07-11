# XLSXWriter

Эта библиотека предназначена для создания xlsx документов с минимальным потреблением памяти.

## Основные особенности:
+ Низкое потребление памяти в отличии от [PHPExcel](https://github.com/PHPOffice/PHPExcel)
+ Обладает более широкими возможностями и позволяет записывать несколько таблиц на одном листе в отличии от [PHP_XLSXWriter](https://github.com/mk-j/PHP_XLSXWriter)

## Возможности:
* Запись огромных таблиц, более 100 000 строк без особого потребления памяти
* Поддержка простых стилей
* Поддержка формул
* Начальная поддерка условного форматирования

## Планируется:
* Простое задание стилей в формате CSS
* Класс-обертка для легкого перехода с PHPExcel на XLSXWriter
* Полноценная поддерка условного форматирования
* Поддерка графиков
* Вставка картинок

## Документация
1. [Создание листов и сохранение документа](#create_sheet)
1. [Запись строки](#write_row)
1. [Стили](#styles)
1. [Условное форматирование](#conditional_formatting)

<a name="create_sheet"></a>
### Создание Листов

```php
/* Создаем документ */
$xlsx = new XLSX();
/*
Создаем новые листы

$sheet_title - название листа
$columns_width - ширина колонок
$freeze_rows и $freeze_columns - позволяет закрепить область прокрутки
*/
$sheet1 = $xlsx->createSheet($sheet_title='Sheet1', $columns_width = array(10,20,30,10,20), $freeze_rows=3, $freeze_columns=2);
$sheet2 = $xlsx->createSheet($sheet_title='Sheet2');

/* Сохраняем документ в файл */
$xlsx->writeToFile('filename.xlsx');
```

<a name="write_row"></a>
### Запись строки
Библиотека рассчитана на запись целой строки, а не ячейки как в PHPExcel. Одновременно со значениеми ячеек можно указать стили либо для всей строки сразу, либо для каждой ячейки в соответствии с ключем либо порядковым номером.

```php
$sheet->writeRow(array('Some text1', 'Some text2'));
/*
Записываем строку c пропуском одной ячейки. Таким образом можно записать несколько таблиц по горизонтали.
$row_data - массив значений
$style - стили ячеек
$startColumn - указывает с какой колонки начинаем писать строку
$row - указывает номер строки, если 0 - запись произойдет в следующую строку
$row_options - массив опций строки, таких как:
	height - высота строки
	hidden - строка скрыта
	collapsed - строка свернута
*/
$sheet->writeRow($row_data = array('Title1','Title2','Title3',null, 'Title5'),
                 $style = array(),
                 $startColumn = 2,
                 $row = 0,
                 $row_options = array());

$row_data = array(null,null,'col1' => 'String value', 'col2' => 10, 'col3' => 20, 'col4' => null, 'col5' => '=D3/E3');
$styles = array
(
	'col5' => ['number_format' => '#0.00%'],
	'col1' => ['fill'=>'#eee', 'halign' => 'right'],
	'col2' => ['color' => '#F00']
);
$sheet->writeRow($row_data, $styles);
```

<a name="styles"></a>
### Стили
Стили задаются ассоциативным массивом.
На данный момент доступны следующие стили:
1.	 Числовой формат
	- **number_format** - числовой формат в формате Excel, например '#0.00%' -  проценты.

2. Шрифты
	* **font** - семейство шрифта
	* **font-size** - высота шрифта
	* **font-style** - стиль шрифта, можно указать следующие возможные значения через пробел или запятую: *bold*, *italic*, *strike*, *underline*
	* **color** - цвет, например #fff, #fe88fe

3. Заливка
	* **fill** - цвет, например #fff, #fe88fe

4. Выравнивание
	* **align** - выравнивание по горизонтали, возможное значение: *general*, *left*, *right*, *justify*, *center*
	* **valign** - выравнивание по вертикали, возможное значение: *bottom*, *center*, *distributed*, *top*
	* **word-wrap** - включает автоперенос текста

5. Границы ячейки
	* **border-color** - цвет, например #fff, #fe88fe
	* **border-style** - стиль границы, возможное значение: *none*, *dashDot*, *dashDotDot*, *dashed*, *dotted*, *double*, *hair*, *medium*, *mediumDashDot*, *mediumDashDotDot*, *mediumDashed*, *slantDashDot*, *thick*, *thin*
	* **border** - указывает стороны к которым будут примененны значение **border-color** и **border-style**. Возможные значения через запятую или пробел: *left*, *right*, *top*, *bottom*.

	Общие значения можно перезаписать для конкртетной стороны границы используя следующие стили:
    * **border-top-color**, **border-top-style**,
    * **border-right-color**, **border-right-style**,
    * **border-bottom-color**, **border-bottom-style**,
    * **border-left-color**, **border-left-style**

<a name="conditional_formatting"></a>
### Условное форматирование
На данный момент реализовано только условное форматирование по формуле. Внимание, не следует создавать условное форматирование для каждой строки! Используйте диапазоны и абсолютные/относительные значения в формуле.

```php
/*
	$range - диапазон ячеек к которому применяется условное форматирование.
	$formula - формула, по условию которой будет применено условное форматирование.
	$styles - стиль условного форматирования. На данный момент доступно изменение заливки, шрифта и границ ячейки.
*/
$sheet->setConditionalStyle($range = 'E1:H1001',$formula = '$A1<=0',$styles = array('color' => '#f00'));
```
