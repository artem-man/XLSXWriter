<?php
/*
 * @Artem Myrhorodskyi
 * */

namespace XLSXWriter\XLSXStyle;

class XLSXFill implements IXLSXStyle
{
	const ALLOWED_STYLE_KEYS = array('fill');

	protected $color = 'FF000000';

	public function __construct(array &$style)
	{
		if (isset($style['fill']) &&
		    $color = XLSXColor::parseColor($style['fill'])) {
				$this->color = $color;
		}
	}

	public function toXML($side='top')
	{
		$xml = '<fill><patternFill patternType="solid"><fgColor rgb="'.strval($this->color).'"/><bgColor rgb="'.strval($this->color).'"/></patternFill></fill>';
		return $xml;
	}
}

