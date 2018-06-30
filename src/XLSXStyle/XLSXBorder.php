<?php
namespace XLSXWriter\XLSXStyle;

/*
 * @Artem Myrhorodskyi
 * */

class XLSXBorder implements IXLSXStyle
{
	const ALLOWED_STYLE_KEYS = array('border',
	                                 'border-color', 'border-style',
	                                 'border-top-color', 'border-top-style',
	                                 'border-right-color', 'border-right-style',
	                                 'border-bottom-color', 'border-bottom-style',
	                                 'border-left-color', 'border-left-style');

	const BORDER_SIDE = array('left', 'right', 'top', 'bottom'); //This order is important :)
	const ALLOWED_BORDER_STYLE = array('none', 'dashDot', 'dashDotDot', 'dashed', 'dotted', 'double','hair','medium','mediumDashDot','mediumDashDotDot','mediumDashed','slantDashDot','thick','thin');

	protected $borders = array();

	public function __construct(array &$style)
	{
		if (isset($style['border'])) {
			$pieces = explode(',', $style['border']);
			foreach ($pieces as $side) {
				$side = trim($side);
				if (!in_array($side, self::BORDER_SIDE)) {
					continue;
				}
				$this->borders[$side] = new XLSXBorderSide($style);
			}
		}
		foreach (self::BORDER_SIDE as $side) {
			$sideStyle = array();
			if (isset($style['border-'.$side.'-style'])) {
				$sideStyle['border-style'] = $style['border-'.$side.'-style'];
			}
			if (isset($style['border-'.$side.'-color'])) {
				$sideStyle['border-color'] = $style['border-'.$side.'-color'];
			}
			if (!empty($sideStyle)) {
				if (isset($this->borders[$side])) {
					$this->borders[$side]->fromArray($sideStyle);
				}
				else {
					$this->borders[$side] = new XLSXBorderSide($sideStyle);
				}
			}
		}
	}

	public function toXML($id=0)
	{
		$xml = '<border diagonalDown="false" diagonalUp="false">';
		foreach (self::BORDER_SIDE as $side)
		{
			if (isset($this->borders[$side])) {
				$xml .= $this->borders[$side]->toXML($side);
			}
			else {
				$xml .= '<' . $side . ' />';
			}
		}
		$xml .= '<diagonal/>';
		$xml .= '</border>';
		return $xml;
	}
}

class XLSXBorderSide implements IXLSXStyle
{
	protected $color;
	protected $style = 'hair';

	public function __construct(array &$style)
	{
		$this->fromArray($style);
	}

	public function fromArray(array &$style)
	{
		if (isset($style['border-style']) && in_array($style['border-style'], self::ALLOWED_BORDER_STYLE)) {
			$this->style = $style['border-style'];
		}
		if (isset($style['border-color']) &&
		    $color = XLSXColor::parseColor($style['border-color'])) {
				$this->color = $color;
		}
	}

	public function toXML($side='top')
	{
		$xml = '<' . $side . ' style="' . $this->style . '">' .
		       ($this->color ? ('<color rgb="'.strval($this->color).'" />'): '') .
		       '</' . $side .'>';
		return $xml;
	}
}
