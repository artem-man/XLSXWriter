<?php
/*
 * @Artem Myrhorodskyi
 * */

namespace XLSXWriter\XLSXStyle;

interface IXLSXStyle
{
	public function __construct(array &$style);
	public function toXML($id);
}