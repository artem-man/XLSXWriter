<?php
/*
 * @Artem Myrhorodskyi
 * */

namespace XLSXWriter\XLSXStyle;


class XLSXCellStyle implements IXLSXStyle
{
	const ALLOWED_STYLE_KEYS = array
	                           (
	                               'num_fmt_id','fill_id','font_id','border_id',
	                               'align','valign','word-wrap', /* XLSXAlignment  */
	                               'default' /* special val for default style */
	                           );

	protected $numFmtId = 0;
	protected $fontId = 0;
	protected $fillId = 0;
	protected $borderId = 0;

	protected $alignment;

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

		$filtered_style = array_intersect_key($style, array_flip(XLSXAlignment::ALLOWED_STYLE_KEYS));
		if (!empty($filtered_style)) {
			$this->alignment = new XLSXAlignment($style);
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
		if (isset($this->alignment)) {
			$xml->writeAttribute('applyAlignment', 1);
			$xml->writeRaw($this->alignment->toXML());
		}

		$xml->endElement();
		return $xml->flush();
	}
}
