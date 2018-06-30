<?php
use XLSXWriter\XLSX;

include_once("../vendor/autoload.php");


$chars = "abcdefghijklmnopqrstuvwxyz0123456789 ";
$s = '';
for($j=0; $j<16192;$j++)
	$s.= $chars[rand()%36];


$xlsx = new XLSX();
$sheet = $xlsx->createSheet('Sheet1', $columns_width = array(10,10,10,10,20,20,20,20));
$sheet->writeRow(array('Column1','Column2','Column3','Column4','Column5','Column6','Column7','Column8'), $common_style = ['halign' => 'center', 'font-style' => 'bold']);
for($i=0; $i<100000; $i++)
{
	$n1 = 5000 - rand()%10000;
	$n2 = 5000 - rand()%10000;
	$n3 = rand(0,10000)/10000;
	$n4 = rand(0,10000)/10000;
	$s1 = substr($s, rand()%4000, rand()%5+5);
	$s2 = substr($s, rand()%8000, rand()%5+5);
	$s3 = substr($s, rand()%12000, rand()%5+5);
	$s4 = substr($s, rand()%16000, rand()%5+5);
    $sheet->writeRow(array($n1, $n2, $n3, $n4, $s1, $s2, $s3, $s4), array([],['number_format' => '$#,##0.00'],[],['number_format' => '#0.00%'],['font'=>'Arial','font-size'=>10,'font-style'=>'bold', 'fill'=>'#eee', 'halign'=>'center', 'border'=>'left,right,top,bottom'],[],[],[]) );
}
$xlsx->writeToFile('xlsx-100k-rows.xlsx');
echo '#'.floor((memory_get_peak_usage())/1024/1024)."MB"."\n";

