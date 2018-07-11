<?php
use XLSXWriter\XLSX;

include_once '../vendor/autoload.php';

$xlsx = new XLSX();
$sheet = $xlsx->createSheet();
$sheet->writeRow(array('Some text'));
$sheet->writeRow($row_data = array('Title1','Title2','Title3',null, 'Title5'), $style = array(), $startColumn = 2, $row = 0, $row_options = array());
$row_data = array(null,null,'col1' => 'String value', 'col2' => 10, 'col3' => 20, 'col4' => null, 'col5' => '=D3/E3');
$styles = array('col5' => ['number_format' => '#0.00%'], 'col1' => ['fill'=>'#eee', 'align' => 'center'], 'col2' => ['color' => '#F00']);
$sheet->writeRow($row_data, $styles);
$xlsx->writeToFile(str_replace('.php', '.xlsx',$_SERVER['argv'][0]));
echo '#'.floor((memory_get_peak_usage())/1024/1024)."MB"."\n";
