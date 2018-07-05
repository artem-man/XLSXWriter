<?php
namespace XLSXWriter\XLSXStyle;

/*
 * @Artem Myrhorodskyi
 * */

class XLSXConditionalStyle implements IXLSXStyle
{
	protected $font;
	protected $fill;
	protected $border;
/*	protected $numFmt;*/

	const ALLOWED_STYLE_KEYS = array('fill',
	                                 'font', 'font-size', 'font-style', 'color',
	                                 'border',
	                                 'border-color', 'border-style',
	                                 'border-top-color', 'border-top-style',
	                                 'border-right-color', 'border-right-style',
	                                 'border-bottom-color', 'border-bottom-style',
	                                 'border-left-color', 'border-left-style');
	function __construct(array &$style)
	{
		$filtered_style = array_intersect_key($style, array_flip(XLSXFont::ALLOWED_STYLE_KEYS));
		if (!empty($filtered_style)) {
			$this->border = new XLSXFont($style);
		}

		$filtered_style = array_intersect_key($style, array_flip(XLSXFill::ALLOWED_STYLE_KEYS));
		if (!empty($filtered_style)) {
			$this->fill = new XLSXFill($style);
		}

		$filtered_style = array_intersect_key($style, array_flip(XLSXBorder::ALLOWED_STYLE_KEYS));
		if (!empty($filtered_style)) {
			$this->border = new XLSXBorder($style);
		}
	}

	function toXML($id = 0)
	{
		$xml = '<dxf>';
		if (isset($this->font)) {
			$xml .= $this->font->toXML($id);
		}
		if (isset($this->fill)) {
			$xml .= $this->fill->toXML($id);
		}
		if (isset($this->border)) {
			$xml .= $this->border->toXML($id);
		}
		$xml .= '</dxf>';
		return $xml;
	}
}
