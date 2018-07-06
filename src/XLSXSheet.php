<?php
/*
 * @Artem Myrhorodskyi
 * */

namespace XLSXWriter;

const EXCEL_2007_MAX_COL=16384;
const EXCEL_2007_MAX_ROW=1048576;

class XLSXSheet
{
	protected $xlsx;
	protected $file_writer;

	protected $tmpfile;

	protected $column_count = 0;
	protected $row_count = 0;

	protected $merge_cells = array();
	protected $conditionalFormatting = array();

	protected $auto_filter = false;
	protected $freeze_rows;
	protected $freeze_columns;

	protected $max_cell_tag_start = 0;
	protected $max_cell_tag_end = 0;
	protected $finalized = false;

	static protected $activeSheet = null;


	public function __construct(XLSX $xlsx, $col_widths=array(), $freeze_rows=false, $freeze_columns=false)
	{
		$this->xlsx = $xlsx;

		$this->tmpfile = new \TempFile();
		$this->file_writer = new \XLSXWriter_BuffererWriter($this->tmpfile->__toString());
		$this->freeze_rows = $freeze_rows;
		$this->freeze_columns = $freeze_columns;

		$tabselected = 'false';
		if (!self::$activeSheet) {
			self::$activeSheet = $this;
			$tabselected = 'true';
		}

		$max_cell = XLSX::cell(EXCEL_2007_MAX_COL-1, EXCEL_2007_MAX_ROW-1);//XFE1048577

		$this->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
		$this->write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
		$this->write(  '<sheetPr filterMode="false">');
		$this->write(    '<pageSetUpPr fitToPage="false"/>');
		$this->write(  '</sheetPr>');
		$this->max_cell_tag_start = $this->file_writer->ftell();
		$this->write('<dimension ref="A1:' . $max_cell . '"/>');
		$this->max_cell_tag_end = $this->file_writer->ftell();
		$this->write(  '<sheetViews>');
		$this->write(    '<sheetView colorId="64" defaultGridColor="true" rightToLeft="false" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="' . $tabselected . '" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">');
		if ($this->freeze_rows && $this->freeze_columns) {
			$this->write(      '<pane ySplit="'.($this->freeze_rows-1).'" xSplit="'.$this->freeze_columns.'" topLeftCell="'.XLSX::cell($this->freeze_columns, $this->freeze_rows).'" activePane="bottomRight" state="frozen"/>');
			$this->write(      '<selection activeCell="'.XLSX::cell(0, $this->freeze_rows).'" activeCellId="0" pane="topRight" sqref="'.XLSX::cell(0, $this->freeze_rows).'"/>');
			$this->write(      '<selection activeCell="'.XLSX::cell($this->freeze_columns, 1).'" activeCellId="0" pane="bottomLeft" sqref="'.XLSX::cell($this->freeze_columns, 1).'"/>');
			$this->write(      '<selection activeCell="'.XLSX::cell($this->freeze_columns, $this->freeze_rows).'" activeCellId="0" pane="bottomRight" sqref="'.XLSX::cell($this->freeze_columns, $this->freeze_rows).'"/>');
		}
		elseif ($this->freeze_rows) {
			$this->write(      '<pane ySplit="'.($this->freeze_rows-1).'" topLeftCell="'.XLSX::cell(0, $this->freeze_rows).'" activePane="bottomLeft" state="frozen"/>');
			$this->write(      '<selection activeCell="'.XLSX::cell(0, $this->freeze_rows).'" activeCellId="0" pane="bottomLeft" sqref="'.XLSX::cell(0, $this->freeze_rows).'"/>');
		}
		elseif ($this->freeze_columns) {
			$this->write(      '<pane xSplit="'.$this->freeze_columns.'" topLeftCell="'.XLSX::cell($this->freeze_columns, 1).'" activePane="topRight" state="frozen"/>');
			$this->write(      '<selection activeCell="'.XLSX::cell($this->freeze_columns, 1).'" activeCellId="0" pane="topRight" sqref="'.XLSX::cell($this->freeze_columns, 1).'"/>');
		}
		else { // not frozen
			$this->write(      '<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
		}
		$this->write(    '</sheetView>');
		$this->write(  '</sheetViews>');
		$this->write(  '<cols>');

		$i=0;
		if (!empty($col_widths)) {
			foreach($col_widths as $column_width) {
				if (isset($column_width)) {
					$this->write(  '<col collapsed="false" hidden="false" max="'.($i+1).'" min="'.($i+1).'" style="0" customWidth="true" width="'.floatval($column_width).'"/>');
				}
				else {
					$this->write(  '<col collapsed="false" hidden="false" max="'.($i+1).'" min="'.($i+1).'" style="0" customWidth="false" width="11.5"/>');
				}
				$i++;
			}
		}
		$this->write(  '<col collapsed="false" hidden="false" max="1024" min="'.($i+1).'" style="0" customWidth="false" width="11.5"/>');
		$this->write(  '</cols>');
		$this->write(  '<sheetData>');
	}

	public function writeRow(array $row_data, array $style = null, $startColumn = 0, $row = 0, array $row_options = null)
	{
		$this->column_count = max($this->column_count, count($row_data) + $startColumn);

		if (!$row) {
			$row = $this->row_count+1;
		}
		if ($this->row_count >= $row) {
			throw new Exception('Current row less then previous');
		}
		$this->row_count = $row;

		if (!empty($row_options)) {
			$ht = isset($row_options['height']) ? floatval($row_options['height']) : 12.1;
			$customHt = isset($row_options['height']) ? true : false;
			$hidden = isset($row_options['hidden']) ? (bool)($row_options['hidden']) : false;
			$collapsed = isset($row_options['collapsed']) ? (bool)($row_options['collapsed']) : false;
			$this->write('<row collapsed="'.($collapsed).'" customFormat="false" customHeight="'.($customHt).'" hidden="'.($hidden).'" ht="'.($ht).'" outlineLevel="0" r="' . $row . '">');
		}
		else {
			$this->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="" ht="12.1" outlineLevel="0" r="' . $row . '">');
		}

		$custom_cell_style = false;
		if (isset($style) && !empty($style) && is_array(reset($style))) {
			$custom_cell_style = true;
		}

		$col = 0;
		foreach ($row_data as $k => $v) {
			if (isset($v)) {
				if ($custom_cell_style) {
					if (isset($style[$k])) {
						$cell_style = $style[$k];
					}
					elseif (isset($style[$col])) {
						$cell_style = $style[$col];
					}
					else {
						$cell_style = null;
					}
				}
				else {
					$cell_style = $style;
				}

				$this->writeCell($col+$startColumn, $row, $v, $cell_style);
			}
			$col++;
		}
		$this->write('</row>');
		return $this;
	}

	protected function writeCell($column, $row, $value, array $style = null)
	{
		$cell_name = XLSX::cell($column, $row);

		$number_format_type = 'n_auto';
		$cell_style_idx = 0;

		if (isset($style) && !empty($style)) {
			$number_format = 'GENERAL';
			if (isset($style['number_format'])) {
				$number_format = XLSX::numberFormatStandardized($style['number_format']);
				$number_format_type = XLSX::determineNumberFormatType($number_format);
			}
			$cell_style_idx = $this->xlsx->addCellStyle($style);
		}

		switch(true) {
		case (!is_scalar($value) || $value === ''): //objects, array, empty
			$this->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'"/>');
			break;
		case (is_string($value) && $value{0} == '='):
			$this->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="s"><f>'.XLSX::xmlspecialchars($value).'</f></c>');
			break;
		default:
			switch($number_format_type) {
			case 'n_date':
				$this->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="n"><v>'.intval(XLSX::convert_date_time($value)).'</v></c>');
				break;
			case 'n_datetime':
				$this->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="n"><v>'.XLSX::convert_date_time($value).'</v></c>');
				break;
			case 'n_numeric':
				$this->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="n"><v>'.XLSX::xmlspecialchars($value).'</v></c>');//int,float,currency
				break;
			case 'n_string':
				$this->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="inlineStr"><is><t>'.XLSX::xmlspecialchars($value).'</t></is></c>');
				break;
			default: //n_auto and try auto-detect
				if (!is_string($value) || $value=='0' || ($value[0]!='0' && ctype_digit($value)) || preg_match("/^\-?(0|[1-9][0-9]*)(\.[0-9]+)?$/", $value)){
					$this->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="n"><v>'.XLSX::xmlspecialchars($value).'</v></c>');//int,float,currency
				}
				else {
					$this->write('<c r="'.$cell_name.'" s="'.$cell_style_idx.'" t="inlineStr"><is><t>'.XLSX::xmlspecialchars($value).'</t></is></c>');
				}
			}
		}
	}

	public function markMergedCell($range)
	{
		$this->merge_cells[] = $range;
		return $this;
	}

	public function setConditionalStyle($range, $formula, array $style)
	{
		$dxfid = $this->xlsx->setConditionalStyle($style);
		$this->conditionalFormatting[] = array
		(
			'range' => $range,
			'formula' => $formula,
			'dxfid' => $dxfid
		);
	}

	public function hasAutoFilter()
	{
		return $this->auto_filter;
	}

	public function setAutoFilter()
	{
		$this->auto_filter = true;
	}


	public function finalizeSheet()
	{
		if ($this->finalized) {
			return;
		}
		$this->write(    '</sheetData>');

		if (!empty($this->conditionalFormatting)) {
			foreach($this->conditionalFormatting as $i => &$rule) {
				$this->write(PHP_EOL);
				$this->write('<conditionalFormatting sqref="' . XLSX::xmlspecialchars($rule['range']) . '">');
				$this->write(PHP_EOL);
				$this->write('<cfRule type="expression" priority="'.($i+1).'" aboveAverage="0" equalAverage="0" bottom="0" percent="0" rank="0" text="" dxfId="'.$rule['dxfid'].'">');
				$this->write(PHP_EOL);
				$this->write('<formula>'.XLSX::xmlspecialchars($rule['formula']).'</formula>');
				$this->write(PHP_EOL);
				$this->write('</cfRule>');
				$this->write('</conditionalFormatting>');
			}
		}

		if (!empty($this->merge_cells)) {
			$this->write(    '<mergeCells>');
			foreach ($this->merge_cells as $range) {
				$this->write(        '<mergeCell ref="' . $range . '"/>');
			}
			$this->write(    '</mergeCells>');
		}

		$max_cell = XLSX::cell($this->column_count - 1, $this->row_count);

		if ($this->auto_filter) {
			$this->write(    '<autoFilter ref="A1:' . $max_cell . '"/>');
		}

		$this->write(    '<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
		$this->write(    '<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');
		$this->write(    '<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>');
		$this->write(    '<headerFooter differentFirst="false" differentOddEven="false">');
		$this->write(        '<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
		$this->write(        '<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
		$this->write(    '</headerFooter>');
		$this->write('</worksheet>');

		$max_cell_tag = '<dimension ref="A1:' . $max_cell . '"/>';
		$padding_length = $this->max_cell_tag_end - $this->max_cell_tag_start - strlen($max_cell_tag);
		$this->file_writer->fseek($this->max_cell_tag_start);
		$this->write($max_cell_tag.str_repeat(" ", $padding_length));
		$this->file_writer->close();
		$this->finalized=true;
		return $this->tmpfile->__toString();
	}

	protected function write($str)
	{
		$this->file_writer->write($str);
	}
}
