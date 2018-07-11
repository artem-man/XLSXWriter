<?php
/*
 * @Artem Myrhorodskyi
 * */

namespace XLSXWriter\XLSXStyle;

class XLSXAlignment implements IXLSXStyle
{
	const ALLOWED_STYLE_KEYS = array('align','valign','word-wrap');
	const ALLOWED_HORISONTAL_STYLE = array('general','left','right','justify','center');
	const ALLOWED_VERTICAL_STYLE = array('bottom','center','distributed','top');

	protected $hAlignment = 'general';
	protected $vAlignment = 'bottom';
	protected $wrapText = false;


	public function __construct(array &$style)
	{
		if (isset($style['align']) && in_array($style['align'], self::ALLOWED_HORISONTAL_STYLE))
		{
			$this->hAlignment = $style['align'];
		}
		if (isset($style['valign']) && in_array($style['valign'], self::ALLOWED_VERTICAL_STYLE))
		{
			$this->vAlignment = $style['valign'];
		}
		if (isset($style['word-wrap']))
		{
			$this->wrapText = (bool)$style['word-wrap'];
		}

	}

	public function toXML($id=0)
	{
		$xml = new \XMLWriter();
		$xml->openMemory();

		$xml->startElement('alignment');
		$xml->writeAttribute('horizontal', $this->hAlignment);
		$xml->writeAttribute('vertical', $this->vAlignment);
		$xml->writeAttribute('textRotation', 0);
		$xml->writeAttribute('wrapText', $this->wrapText ? 'true': 'false');
		$xml->writeAttribute('indent', 0);
		$xml->writeAttribute('shrinkToFit', 'false');
		$xml->endElement();

		return $xml->flush();
	}
}
