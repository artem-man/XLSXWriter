<?php
/*
 * @Artem Myrhorodskyi
 * */

namespace XLSXWriter\XLSXStyle;

class XLSXFont implements IXLSXStyle
{
	protected $size = 10;
	protected $family = 2;
	protected $name = 'Arial';
	protected $bold = false;
	protected $italic = false;
	protected $strike = false;
	protected $underline = false;
	protected $color;

	const ALLOWED_STYLE_KEYS = array('font', 'font-size', 'font-style', 'color');
	public function __construct(array &$style)
	{
		if (isset($style['font-size']))	{
			$this->size = floatval($style['font-size']);
		}

		if (isset($style['font']) && is_string($style['font']))	{
			switch($style['font']) {
			case 'Times New Roman':
				$this->family = 1;
				break;
			case 'Courier New':
				$this->family = 3;
				break;
			case 'Comic Sans MS':
				$this->family = 4;
				break;
			}
			$this->name = strval($style['font']);
		}

		if (isset($style['font-style']) && is_string($style['font-style'])) {
			if (strpos($style['font-style'], 'bold')!==false) { $this->bold = true; }
			if (strpos($style['font-style'], 'italic')!==false) { $this->italic = true; }
			if (strpos($style['font-style'], 'strike')!==false) { $this->strike = true; }
			if (strpos($style['font-style'], 'underline')!==false) { $this->underline = true; }
		}

		if (isset($style['color']) &&
		    $color = XLSXColor::parseColor($style['color'])) {
			$this->color = $color;
		}
	}

	public function toXML($id=0)
	{
		$xml = '<font>' .
		       '<name val="'.htmlspecialchars($this->name).'"/>'.
		       '<charset val="1"/>'.
		       '<family val="'. $this->family . '"/>' .
		       '<sz val="'. $this->size .'"/>' .
	           (empty($this->color) ? '' : '<color rgb="'.strval($this->color).'"/>') .
		       ($this->bold ? '<b val="true"/>' : '') .
		       ($this->italic ? '<i val="true"/>' : '') .
		       ($this->underline ? '<u val="single"/>' : '') .
		       ($this->strike ? '<strike val="true"/>' : '') .
		       '</font>';
		return $xml;
	}
}
