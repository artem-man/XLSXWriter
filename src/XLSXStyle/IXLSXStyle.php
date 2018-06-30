<?php
namespace XLSXWriter\XLSXStyle;

/*
 * @Artem Myrhorodskyi
 * */

interface IXLSXStyle
{
	public function __construct(array &$style);
	public function toXML($id);
}