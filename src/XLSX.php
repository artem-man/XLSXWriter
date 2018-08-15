<?php
/*
 * @Artem Myrhorodskyi
 * */

namespace XLSXWriter;

class XLSX
{
    const CONDITION_CELLIS       = 'cellIs';
    const CONDITION_CONTAINSTEXT = 'containsText';
    const CONDITION_EXPRESSION   = 'expression';


	//http://www.ecma-international.org/publications/standards/Ecma-376.htm
	//http://officeopenxml.com/SSstyles.php
	//------------------------------------------------------------------
	//http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
	//------------------------------------------------------------------
	protected $title;
	protected $subject;
	protected $author;
	protected $company;
	protected $description;
	protected $keywords = array();

	protected $sheets = array();
	protected $xlsxstyle = null;

	public function __construct($defaultStyle = array())
	{
		$this->xlsxstyle = new XLSXStyle($defaultStyle);
	}

	public function createSheet($sheet_name = '', $col_widths=array(), $freeze_rows=false, $freeze_columns=false )
	{
		if (empty($sheet_name)) {
			$sheet_name = 'Sheet'. (count($this->sheets)+1);
		}
		if (isset($this->sheets[$sheet_name])) {
			return $this->sheets[$sheet_name];
		}

		$this->sheets[$sheet_name] = new XLSXSheet($this, $col_widths, $freeze_rows, $freeze_columns);

		return $this->sheets[$sheet_name];
	}

	public function addCellStyle(array $style=array())
	{
//		$style += array('fill' => '#EEECE1', 'color' => '#776F45');
		return $this->xlsxstyle->addCellStyle($style);
	}

	public function setConditionalStyle(array &$style)
	{
		$dxfid = $this->xlsxstyle->addConditionalStyle($style);
		return $dxfid;
	}


	public function writeToStdOut()
	{
		$temp_file = \TempFileCreator::tempFilename();
		$this->writeToFile($temp_file);
		readfile($temp_file);
	}

	public function writeToString()
	{
		$temp_file = \TempFileCreator::tempFilename();
		$this->writeToFile($temp_file);
		$string = file_get_contents($temp_file);
		return $string;
	}

	public function writeToFile($filename)
	{
		if ( file_exists( $filename ) ) {
			if ( is_writable( $filename ) ) {
				@unlink( $filename ); //if the zip already exists, remove it
			} else {
				self::log( "Error in " . __CLASS__ . "::" . __FUNCTION__ . ", file is not writeable." );
				return;
			}
		}
		$zip = new \ZipArchive();
		if (empty($this->sheets))                       { self::log("Error in ".__CLASS__."::".__FUNCTION__.", no worksheets defined."); return; }
		if (!$zip->open($filename, \ZipArchive::CREATE)) { self::log("Error in ".__CLASS__."::".__FUNCTION__.", unable to create zip."); return; }

		$zip->addEmptyDir("docProps/");
		$zip->addFromString("docProps/app.xml" , $this->buildAppXML() );
		$zip->addFromString("docProps/core.xml", $this->buildCoreXML());

		$zip->addEmptyDir("_rels/");
		$zip->addFromString("_rels/.rels", $this->buildRelationshipsXML());

		$zip->addEmptyDir("xl/worksheets/");
		$i = 0;
		foreach($this->sheets as $sheet) {
			$zip->addFile($sheet->finalizeSheet(), "xl/worksheets/". self::getXMLNameForWorksheetNum($i));
			$i++;
		}
		$zip->addFromString("xl/workbook.xml"         , $this->buildWorkbookXML() );
		$zip->addFile($this->xlsxstyle->writeStylesXML(), "xl/styles.xml" );
		$zip->addFromString("[Content_Types].xml"     , $this->buildContentTypesXML() );

		$zip->addEmptyDir("xl/_rels/");
		$zip->addFromString("xl/_rels/workbook.xml.rels", $this->buildWorkbookRelsXML() );
		$zip->close();
	}

	protected function buildAppXML()
	{
		$app_xml="";
		$app_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$app_xml.='<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
		$app_xml.='<TotalTime>0</TotalTime>';
		$app_xml.='<Company>'.self::xmlspecialchars($this->company).'</Company>';
		$app_xml.='</Properties>';
		return $app_xml;
	}

	protected function buildCoreXML()
	{
		$core_xml="";
		$core_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$core_xml.='<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
		$core_xml.='<dcterms:created xsi:type="dcterms:W3CDTF">'.date("Y-m-d\TH:i:s.00\Z").'</dcterms:created>';//$date_time = '2014-10-25T15:54:37.00Z';
		$core_xml.='<dc:title>'.self::xmlspecialchars($this->title).'</dc:title>';
		$core_xml.='<dc:subject>'.self::xmlspecialchars($this->subject).'</dc:subject>';
		$core_xml.='<dc:creator>'.self::xmlspecialchars($this->author).'</dc:creator>';
		if (!empty($this->keywords)) {
			$core_xml.='<cp:keywords>'.self::xmlspecialchars(implode (", ", (array)$this->keywords)).'</cp:keywords>';
		}
		$core_xml.='<dc:description>'.self::xmlspecialchars($this->description).'</dc:description>';
		$core_xml.='<cp:revision>0</cp:revision>';
		$core_xml.='</cp:coreProperties>';
		return $core_xml;
	}

	protected function buildRelationshipsXML()
	{
		$rels_xml="";
		$rels_xml.='<?xml version="1.0" encoding="UTF-8"?>'."\n";
		$rels_xml.='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
		$rels_xml.='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
		$rels_xml.='<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
		$rels_xml.='<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
		$rels_xml.="\n";
		$rels_xml.='</Relationships>';
		return $rels_xml;
	}
//MAN
	protected function buildWorkbookXML()
	{
		$i=0;
		$workbook_xml="";
		$workbook_xml.='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'."\n";
		$workbook_xml.='<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
		$workbook_xml.='<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
		$workbook_xml.='<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
		$workbook_xml.='<sheets>';
		foreach($this->sheets as $sheet_name=>$sheet) {
			$sheetname = self::sanitize_sheetname($sheet_name);
			$workbook_xml.='<sheet name="'.self::xmlspecialchars($sheetname).'" sheetId="'.($i+1).'" state="visible" r:id="rId'.($i+2).'"/>';
			$i++;
		}
		$workbook_xml.='</sheets>';
		$workbook_xml.='<definedNames>';
		foreach($this->sheets as $sheet_name=>$sheet) {
			if ($sheet->hasAutoFilter()) {
				$sheetname = self::sanitize_sheetname($sheet_name);
				$workbook_xml.='<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\''.self::xmlspecialchars($sheetname).'\'!$A$1:' . self::cell($sheet->column_count - 1, $sheet->row_count, true) . '</definedName>';
				$i++;
			}
		}
		$workbook_xml.='</definedNames>';
		$workbook_xml.='<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';
		return $workbook_xml;
	}

	protected function buildWorkbookRelsXML()
	{
		$wkbkrels_xml="";
		$wkbkrels_xml.='<?xml version="1.0" encoding="UTF-8"?>'."\n";
		$wkbkrels_xml.='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
		$wkbkrels_xml.='<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
		$i=0;
		foreach($this->sheets as $sheet_name=>$sheet) {
			$wkbkrels_xml.='<Relationship Id="rId'.($i+2).'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/'.self::getXMLNameForWorksheetNum($i).'"/>';
			$i++;
		}
		$wkbkrels_xml.="\n";
		$wkbkrels_xml.='</Relationships>';
		return $wkbkrels_xml;
	}

	protected function buildContentTypesXML()
	{
		$content_types_xml="";
		$content_types_xml.='<?xml version="1.0" encoding="UTF-8"?>'."\n";
		$content_types_xml.='<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
		$content_types_xml.='<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
		$content_types_xml.='<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
		$i = 0;
		foreach($this->sheets as $sheet_name=>$sheet) {
			$content_types_xml.='<Override PartName="/xl/worksheets/'. self::getXMLNameForWorksheetNum($i).'" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
			$i++;
		}
		$content_types_xml.='<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
		$content_types_xml.='<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
		$content_types_xml.='<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
		$content_types_xml.='<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
		$content_types_xml.="\n";
		$content_types_xml.='</Types>';
		return $content_types_xml;
	}

	protected static function getXMLNameForWorksheetNum($i)
	{
		return 'sheet' . ($i+1) .".xml";
	}

	public function setTitle($title='')
	{
		$this->title=$title;
	}

	public function setSubject($subject='')
	{
		$this->subject=$subject;
	}

	public function setAuthor($author='')
	{
		$this->author=$author;
	}

	public function setCompany($company='')
	{
		$this->company=$company;
	}

	public function setKeywords($keywords='')
	{
		$this->keywords=$keywords;
	}

	public function setDescription($description='')
	{
		$this->description=$description;
	}


	//------------------------------------------------------------------
	/*
	 * @param $column_number int, zero based
	 * @param $row_number int, zero based
	 * @param $absolute bool
	 * @return Cell label/coordinates, ex: A1, C3, AA42 (or if $absolute==true: $A$1, $C$3, $AA$42)
	 * */
	public static function cell($column_number, $row_number, $absolute=false)
	{
		$n = $column_number;
		for($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
			$r = chr($n%26 + 0x41) . $r;
		}
		if ($absolute) {
			return '$' . $r . '$' . $row_number;
		}
		return $r . $row_number;
	}


	//------------------------------------------------------------------
	public static function log($string)
	{
		file_put_contents("php://stderr", date("Y-m-d H:i:s:").rtrim(is_array($string) ? json_encode($string) : $string)."\n");
	}
	//------------------------------------------------------------------
	//------------------------------------------------------------------
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
		case (preg_match("/$/", $num_format)):
		case (preg_match("/%/", $num_format)):
		case (preg_match("/0/", $num_format)):
			return 'n_numeric';
		case (preg_match("/[H]{1,2}:[M]{1,2}/i", $num_format)):
		case (preg_match("/[M]{1,2}:[S]{1,2}/i", $num_format)):
			 return 'n_datetime';
		case (preg_match("/[Y]{2,4}/i", $num_format)):
		case (preg_match("/[D]{1,2}/i", $num_format)):
		case (preg_match("/[M]{1,2}/i", $num_format)):
			return 'n_date';
		}
		return 'n_auto';
	}
	//------------------------------------------------------------------
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
	//------------------------------------------------------------------
	public static function convert_date_time($date_input)
	{
		$days    = 0;    // Number of days since epoch
		$seconds = 0;    // Time expressed as fraction of 24h hours in seconds
		$year=$month=$day=0;
		$hour=$min  =$sec=0;

		$date_time = $date_input;
		if (preg_match("/(\d{4})\-(\d{2})\-(\d{2})/", $date_time, $matches))
		{
			list($junk,$year,$month,$day) = $matches;
		}
		if (preg_match("/(\d+):(\d{2}):(\d{2})/", $date_time, $matches))
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

	//------------------------------------------------------------------
	public static function sanitize_filename($filename) //http://msdn.microsoft.com/en-us/library/aa365247%28VS.85%29.aspx
	{
		$nonprinting = array_map('chr', range(0,31));
		$invalid_chars = array('<', '>', '?', '"', ':', '|', '\\', '/', '*', '&');
		$all_invalids = array_merge($nonprinting,$invalid_chars);
		return str_replace($all_invalids, "", $filename);
	}
	//------------------------------------------------------------------
	public static function sanitize_sheetname($sheetname)
	{
		static $badchars  = '\\/?*:[]';
		static $goodchars = '        ';
		$sheetname = strtr($sheetname, $badchars, $goodchars);
		$sheetname = substr($sheetname, 0, 31);
		$sheetname = trim(trim(trim($sheetname),"'"));//trim before and after trimming single quotes
		return !empty($sheetname) ? $sheetname : 'Sheet'.((rand()%900)+100);
	}
	//------------------------------------------------------------------
	public static function xmlspecialchars($val)
	{
		//note, badchars does not include \t\n\r (\x09\x0a\x0d)
		static $badchars = "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x7f";
		static $goodchars = "                              ";
		return strtr(htmlspecialchars($val, ENT_QUOTES | ENT_XML1), $badchars, $goodchars);//strtr appears to be faster than str_replace
	}
}
