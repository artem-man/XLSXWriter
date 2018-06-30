<?php
namespace XLSXWriter\XLSXStyle;

/*
 * @Artem Myrhorodskyi
 * */

class XLSXNumberFormat implements IXLSXStyle
{
	protected $number_format = 'GENERAL';

	const ALLOWED_STYLE_KEYS = array('number_format');

	function __construct(array &$style)
	{
		if (isset($style['number_format'])) {
			$this->number_format = self::numberFormatStandardized($style['number_format']);
		}
	}

	function toXML($id)
	{
		$xml = '<numFmt numFmtId="'. $id .'" formatCode="'.\XLSXWriter\XLSX::xmlspecialchars($this->number_format).'" />';
		return $xml;
	}

	public static function numberFormatStandardized($num_format)
	{
		if ($num_format=='money') { $num_format='dollar'; }
		if ($num_format=='number') { $num_format='integer'; }

		if      ($num_format=='string')   $num_format='@';
		else if ($num_format=='integer')  $num_format='0';
		else if ($num_format=='date')     $num_format='YYYY-MM-DD';
		else if ($num_format=='datetime') $num_format='YYYY-MM-DD HH:MM:SS';
		else if ($num_format=='price')    $num_format='#,##0.00';
		else if ($num_format=='dollar')   $num_format='[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00';
		else if ($num_format=='euro')     $num_format='#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]';
		$ignore_until='';
		$escaped = '';
		for($i=0,$ix=strlen($num_format); $i<$ix; $i++)
		{
			$c = $num_format[$i];
			if ($ignore_until=='' && $c=='[')
				$ignore_until=']';
			else if ($ignore_until=='' && $c=='"')
				$ignore_until='"';
			else if ($ignore_until==$c)
				$ignore_until='';
			if ($ignore_until=='' && ($c==' ' || $c=='-'  || $c=='('  || $c==')') && ($i==0 || $num_format[$i-1]!='_'))
				$escaped.= "\\".$c;
			else
				$escaped.= $c;
		}
		return $escaped;
	}
}
