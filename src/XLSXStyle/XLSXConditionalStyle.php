<?php
/*
 * @Artem Myrhorodskyi
 * */

namespace XLSXWriter\XLSXStyle;

class XLSXConditionalStyle implements IXLSXStyle
{
	protected $font;
	protected $fill;
	protected $border;
	protected $alignment;
/*
	Num format not implemented at this time
	Support of num format require add same rows to <dxf> and to <numFmts>
*/

/*
	protected $numFmt;
*/

	const ALLOWED_STYLE_KEYS = array('fill', /* XLSXFill */
	                                 'font', 'font-size', 'font-style', 'color', /* XLSXFont */
	                                 'align','valign','word-wrap', /* XLSXAligment */
	                                 'number_format',
	                                 'border',
	                                 'border-color', 'border-style',
	                                 'border-top-color', 'border-top-style',
	                                 'border-right-color', 'border-right-style',
	                                 'border-bottom-color', 'border-bottom-style',
	                                 'border-left-color', 'border-left-style');
	function __construct(array &$style)
	{
		$filtered_style = array_intersect_key($style, array_flip(XLSXFill::ALLOWED_STYLE_KEYS));
		if (!empty($filtered_style)) {
			$this->fill = new XLSXFill($style);
		}

		$filtered_style = array_intersect_key($style, array_flip(XLSXFont::ALLOWED_STYLE_KEYS));
		if (!empty($filtered_style)) {
			$this->border = new XLSXFont($style);
		}

		$filtered_style = array_intersect_key($style, array_flip(XLSXAlignment::ALLOWED_STYLE_KEYS));
		if (!empty($filtered_style)) {
			$this->alignment = new XLSXAlignment($style);
		}

		$filtered_style = array_intersect_key($style, array_flip(XLSXBorder::ALLOWED_STYLE_KEYS));
		if (!empty($filtered_style)) {
			$this->border = new XLSXBorder($style);
		}

/*		$filtered_style = array_intersect_key($style, array_flip(XLSXNumberFormat::ALLOWED_STYLE_KEYS));
		if (!empty($filtered_style)) {
			$this->numFmt = new XLSXNumberFormat($style);
		}*/
	}

	function toXML($id = 0)
	{
		$xml = '<dxf>';
		if (isset($this->fill)) {
			$xml .= $this->fill->toXML($id);
		}

		if (isset($this->font)) {
			$xml .= $this->font->toXML($id);
		}

		if (isset($this->alignment)) {
			$xml .= $this->alignment->toXML($id);
		}

		if (isset($this->border)) {
			$xml .= $this->border->toXML($id);
		}

/*		if (isset($this->numFmt)) {
			$xml .= $this->numFmt->toXML($id+164);
		}*/
		$xml .= '</dxf>';
		return $xml;
	}
}
