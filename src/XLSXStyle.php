<?php
/*
 * @Artem Myrhorodskyi
 * */

namespace XLSXWriter;

use XLSXWriter\XLSXStyle\XLSXCellStyle;
use XLSXWriter\XLSXStyle\XLSXNumberFormat;
use XLSXWriter\XLSXStyle\XLSXFont;
use XLSXWriter\XLSXStyle\XLSXFill;
use XLSXWriter\XLSXStyle\XLSXBorder;
use XLSXWriter\XLSXStyle\XLSXConditionalStyle;

class XLSXStyle
{
	protected $cell_styles = array();
	protected $defaultStyle = array();

	protected $styleIdx;
	protected $fontsIdx;
	protected $numberFormatsIdx;
	protected $bordersIdx;
	protected $fillIdx;

	protected $conditionalStyleIdx;

	public static function add_to_list_get_index(&$haystack, $needle)
	{
		$existing_idx = array_search($needle, $haystack, $strict=true);
		if ($existing_idx === false)
		{
			$existing_idx = count($haystack);
			$haystack[] = $needle;
		}
		return $existing_idx;
	}

	public function __construct($defaultStyle = array())
	{
		$this->defaultStyle = $defaultStyle;

		$this->styleIdx = new \ArrayToObjectsUIndex(XLSXCellStyle::ALLOWED_STYLE_KEYS);
		$this->numberFormatsIdx = new \ArrayToObjectsUIndex(XLSXNumberFormat::ALLOWED_STYLE_KEYS, 164);
		$this->fontsIdx = new \ArrayToObjectsUIndex(XLSXFont::ALLOWED_STYLE_KEYS, 4);
		$this->bordersIdx = new \ArrayToObjectsUIndex(XLSXBorder::ALLOWED_STYLE_KEYS, 1);
		$this->fillIdx = new \ArrayToObjectsUIndex(XLSXFill::ALLOWED_STYLE_KEYS, 2);
		$this->conditionalStyleIdx = new \ArrayToObjectsUIndex(XLSXConditionalStyle::ALLOWED_STYLE_KEYS);

		$defaultStyle['default'] = true;  //Set default cell style
		$this->addCellStyle($defaultStyle);
	}

	public function addConditionalStyle(array &$style)
	{
		$id = $this->conditionalStyleIdx->lookup($style,
			function($filtered_data) {
				return new XLSXConditionalStyle($filtered_data);
			}
		);
		return $id;
	}

	public function addCellStyle(array &$style)
	{
		$style += $this->defaultStyle;

		$style_indexes = array();

		$id = $this->numberFormatsIdx->lookup($style,
			function($filtered_data) {
				return new XLSXNumberFormat($filtered_data);
			}
		);
		if ($id !== false) {
			$style_indexes['num_fmt_id'] = $id;
		}

		$id = $this->fillIdx->lookup($style,
			function($filtered_data) {
				return new XLSXFill($filtered_data);
			}
		);
		if ($id !== false) {
			$style_indexes['fill_id'] = $id;
		}

		$id = $this->fontsIdx->lookup($style,
			function($filtered_data) {
				return new XLSXFont($filtered_data);
			}
		);
		if ($id !== false) {
			$style_indexes['font_id'] = $id;
		}

		$id = $this->bordersIdx->lookup($style,
			function($filtered_data) {
				return new XLSXBorder($filtered_data);
			}
		);
		if ($id !== false) {
			$style_indexes['border_id'] = $id;
		}

		$style_indexes += $style;
		$id = $this->styleIdx->lookup(($style_indexes),
			function($filtered_data) {
				return new XLSXCellStyle($filtered_data);
			}
		);
		return $id;
	}


	public function writeStylesXML()
	{
		$temporary_filename = \TempFileCreator::tempFilename();
		$file = new \XLSXWriter_BuffererWriter($temporary_filename);
		$file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n");
		$file->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');

		//Write numbers format
		$file->write('<numFmts count="'. ($this->numberFormatsIdx->count() - 164) .'">');
		foreach($this->numberFormatsIdx->data() as $id => &$obNumberFormat) {
			$file->write($obNumberFormat->toXML($id));
		}
		$file->write('</numFmts>');

		//Write fonts
		$file->write('<fonts count="' . $this->fontsIdx->count() . '">');
		$file->write(		'<font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font>');
		$file->write(		'<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
		$file->write(		'<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
		$file->write(		'<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
		foreach($this->fontsIdx->data() as &$font) {
			$file->write($font->toXML());
		}
		$file->write('</fonts>');

		//Write fills
		$file->write('<fills count="'. $this->fillIdx->count().'">');
		$file->write(	'<fill><patternFill patternType="none"/></fill>');
		$file->write(	'<fill><patternFill patternType="gray125"/></fill>');
		foreach($this->fillIdx->data() as &$fill) {
			$file->write($fill->toXML());
		}
		$file->write('</fills>');

		//Write borders
		$file->write('<borders count="'.$this->bordersIdx->count().'">');
        $file->write(    '<border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border>');
		foreach($this->bordersIdx->data() as &$border) {
			$file->write($border->toXML());
		}
		$file->write('</borders>');

		$file->write('<cellStyleXfs count="1">');
		$file->write(	'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" />');
		$file->write('</cellStyleXfs>');

		$file->write('<cellXfs count="'.$this->styleIdx->count() .'">');
		foreach($this->styleIdx->data() as &$style) {
			$file->write($style->toXML());
		}
		$file->write('</cellXfs>');

		$file->write('<cellStyles count="1">');
		$file->write(	'<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>');
		$file->write('</cellStyles>');

		if ($this->conditionalStyleIdx->count()) {
			$file->write('<dxfs count="'.$this->conditionalStyleIdx->count().'">');
			foreach ($this->conditionalStyleIdx->data() as &$style) {
				$file->write($style->toXML());
			}
			$file->write('</dxfs>');
		}

		$file->write('</styleSheet>');
		$file->close();
		return $temporary_filename;
	}
}
