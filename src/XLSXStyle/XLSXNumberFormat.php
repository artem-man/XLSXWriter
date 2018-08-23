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

	public static function determineNumberFormatType($num_format)
	{
		$num_format = preg_replace("/\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)\]/i", "", $num_format);
		switch (true)
		{
		case ($num_format=='GENERAL'):
			 return 'n_auto';
		case ($num_format=='@'):
			 return 'n_string';
		case ($num_format=='0'):
			 return 'n_numeric';
		case (preg_match('/\$/', $num_format)):
		case (preg_match('/%/', $num_format)):
		case (preg_match('/0/', $num_format)):
			return 'n_numeric';
		case (preg_match('/[H]{1,2}:[M]{1,2}/i', $num_format)):
		case (preg_match('/[M]{1,2}:[S]{1,2}/i', $num_format)):
			 return 'n_datetime';
		case (preg_match('/[Y]{2,4}/i', $num_format)):
		case (preg_match('/[D]{1,2}/i', $num_format)):
		case (preg_match('/[M]{1,2}/i', $num_format)):
			return 'n_date';
		}
		return 'n_auto';
	}
	//------------------------------------------------------------------
	public static function convert_date_time($date_input)
	{
		$days    = 0;    // Number of days since epoch
		$seconds = 0;    // Time expressed as fraction of 24h hours in seconds
		$year=$month=$day=0;
		$hour=$min  =$sec=0;

		$date_time = $date_input;
		if (preg_match('/(\d{4})\-(\d{2})\-(\d{2})/', $date_time, $matches))
		{
			list($junk,$year,$month,$day) = $matches;
		}
		if (preg_match('/(\d+):(\d{2}):(\d{2})/', $date_time, $matches))
		{
			list($junk,$hour,$min,$sec) = $matches;
			$seconds = ( $hour * 60 * 60 + $min * 60 + $sec ) / ( 24 * 60 * 60 );
		}

		//using 1900 as epoch, not 1904, ignoring 1904 special case

		// Special cases for Excel.
		if ("$year-$month-$day"=='1899-12-31')  return $seconds      ;    // Excel 1900 epoch
		if ("$year-$month-$day"=='1900-01-00')  return $seconds      ;    // Excel 1900 epoch
		if ("$year-$month-$day"=='1900-02-29')  return 60 + $seconds ;    // Excel false leapday

		// We calculate the date by calculating the number of days since the epoch
		// and adjust for the number of leap days. We calculate the number of leap
		// days by normalising the year in relation to the epoch. Thus the year 2000
		// becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
		$epoch  = 1900;
		$offset = 0;
		$norm   = 300;
		$range  = $year - $epoch;

		// Set month days and check for leap year.
		$leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100)) ) ? 1 : 0;
		$mdays = array( 31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 );

		// Some boundary checks
		if($year < $epoch || $year > 9999) return 0;
		if($month < 1     || $month > 12)  return 0;
		if($day < 1       || $day > $mdays[ $month - 1 ]) return 0;

		// Accumulate the number of days since the epoch.
		$days = $day;    // Add days for current month
		$days += array_sum( array_slice($mdays, 0, $month-1 ) );    // Add days for past months
		$days += $range * 365;                      // Add days for past years
		$days += intval( ( $range ) / 4 );             // Add leapdays
		$days -= intval( ( $range + $offset ) / 100 ); // Subtract 100 year leapdays
		$days += intval( ( $range + $offset + $norm ) / 400 );  // Add 400 year leapdays
		$days -= $leap;                                      // Already counted above

		// Adjust for Excel erroneously treating 1900 as a leap year.
		if ($days > 59) { $days++;}

		return $days + $seconds;
	}

}
