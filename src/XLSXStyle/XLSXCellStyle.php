<?php
namespace XLSXWriter\XLSXStyle;

/*
 * @Artem Myrhorodskyi
 * */

class XLSXCellStyle implements IXLSXStyle
{
	const ALLOWED_STYLE_KEYS = array('num_fmt_id','fill_id','font_id','border_id','halign','valign','word-wrap', 'default');
	const ALLOWED_HORISONTAL_STYLE = array('general','left','right','justify','center');
	const ALLOWED_VERTICAL_STYLE = array('bottom','center','distributed','top');

	protected $applyAlignment = false;
	protected $hAlignment = 'general';
	protected $vAlignment = 'bottom';
	protected $wrapText = false;

	protected $numFmtId = 0;
	protected $fontId = 0;
	protected $fillId = 0;
	protected $borderId = 0;

	public function __construct(array &$style)
	{
		if (isset($style['num_fmt_id'])) {
			$this->numFmtId = $style['num_fmt_id'];
		}
		if (isset($style['fill_id'])) {
			$this->fillId = $style['fill_id'];
		}
		if (isset($style['font_id'])) {
			$this->fontId = $style['font_id'];
		}
		if (isset($style['border_id'])) {
			$this->borderId = $style['border_id'];
		}

		if (isset($style['halign']) && in_array($style['halign'], self::ALLOWED_HORISONTAL_STYLE))
		{
			$this->applyAlignment = true;
			$this->hAlignment = $style['halign'];
		}
		if (isset($style['valign']) && in_array($style['valign'], self::ALLOWED_VERTICAL_STYLE))
		{
			$this->applyAlignment = true;
			$this->vAlignment = $style['valign'];
		}
		if (isset($style['word-wrap']))
		{
			$this->applyAlignment = true;
			$this->wrapText = (bool)$style['word-wrap'];
		}

	}

	public function toXML($id=0)
	{
		$xml = new \XMLWriter();
		$xml->openMemory();
		$xml->startElement('xf');
		$xml->writeAttribute('numFmtId', $this->numFmtId);
		$xml->writeAttribute('fontId', $this->fontId);
		$xml->writeAttribute('fillId', $this->fillId);
		$xml->writeAttribute('borderId', $this->borderId);
		if ($this->numFmtId > 0) {
			$xml->writeAttribute('applyNumberFormat', 1);
		}
		if ($this->fontId > 0) {
			$xml->writeAttribute('applyFont', 1);
		}
		if ($this->fillId > 0) {
			$xml->writeAttribute('applyFill', 1);
		}
		if ($this->borderId > 0) {
			$xml->writeAttribute('applyBorder', 1);
		}
		$xml->writeAttribute('xfId', 0);
		if ($this->applyAlignment) {
			$xml->writeAttribute('applyAlignment', 1);
			$xml->startElement('alignment');
			$xml->writeAttribute('horizontal', $this->hAlignment);
			$xml->writeAttribute('vertical', $this->vAlignment);
			$xml->writeAttribute('textRotation', 0);
			$xml->writeAttribute('wrapText', $this->wrapText ? 'true': 'false');
			$xml->writeAttribute('indent', 0);
			$xml->writeAttribute('shrinkToFit', 'false');
			$xml->endElement();
		}

		$xml->endElement();
		return $xml->flush();
	}
}
