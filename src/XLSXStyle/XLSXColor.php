<?php
/*
 * @Artem Myrhorodskyi
 * */

namespace XLSXWriter\XLSXStyle;

class XLSXColor
{
	static public function parseColor($color)
	{
		if (is_string($color) && ($color[0]  == '#')) {
			$v = substr($color,1,6);
			$v = strlen($v)==3 ? $v[0].$v[0].$v[1].$v[1].$v[2].$v[2] : $v;// expand cf0 => ccff00
			return 'FF'.strtoupper($v);
		}
		return false;
	}
}
