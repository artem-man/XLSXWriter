<?php
/*
 * @Artem Myrhorodskyi
 * */

namespace XLSXWriter\XLSXStyle;

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

	public function toXML($id)
	{
		$xml = '<numFmt numFmtId="'. $id .'" formatCode="'.\XLSXWriter\XLSX::xmlspecialchars($this->number_format).'" />';
		return $xml;
	}

	public static function numberFormatStandardized($num_format)
	{
		switch ($num_format) {
		case 'string':
			$num_format = '@';
			break;
		case 'price':
			$num_format = '#,##0.00';
			break;
		case 'money':
		case 'dollar':
			$num_format = '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00';
			break;
		case 'euro':
			$num_format = '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]';
			break;
		case 'integer':
		case 'number':
			$num_format='0';
			break;
		case 'date':
			$num_format='YYYY-MM-DD';
			break;
		case 'datetime':
			$num_format='YYYY-MM-DD HH:MM:SS';
			break;
		}

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
