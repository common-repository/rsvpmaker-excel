<?php
/*
Plugin Name: RSVPMaker Excel
Plugin URI: http://www.rsvpmaker.com
Description: An extension to the RSVPMaker event scheduling plugin for exporting RSVP Reports to Excel.
Author: David F. Carr
Version: 1.1
Author URI: http://www.carrcommunications.com
*/

if(file_exists(WP_PLUGIN_DIR."/rsvpmaker-custom.php") )
	include_once WP_PLUGIN_DIR."/rsvpmaker-custom.php";

$phpexcel_enabled = true;

// helper function for rsvp_excel
function col2chr($a){ 
$a++;
        if($a<27){ 
            return strtoupper(chr($a+96));    
        }else{ 
            while($a > 26){ 
                $b++; 
                $a = $a-26;                
            }                   
            $b = strtoupper(chr($b+96));    
            $a = strtoupper(chr($a+96));                
            return $b.$a; 
        } 
    }

if(!function_exists('rsvp_excel')){

function rsvp_excel() {
if(!isset($_GET["rsvpexcel"]))
	return;
if ( !wp_verify_nonce($_GET['rsvpexcel'],'rsvpexcel') )
{
   print 'Sorry, your nonce did not verify.';
   exit;
}
global $wpdb;
$fields = $_GET["fields"];
$eventid = (int) $_GET["event"];
$columnalpha = array('A',);

include WP_PLUGIN_DIR.'/rsvpmaker-excel/phpexcel/PHPExcel.php'; // include PHP Excel library
	
	$sql = "SELECT post_title FROM ".$wpdb->posts." WHERE ID = $eventid";
	$title = $wpdb->get_var($sql);

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

// Set properties
$objPHPExcel->getProperties()->setCreator("RSVPMaker")
							 ->setLastModifiedBy("RSVPMaker")
							 ->setTitle($title);

$styleArray = array(
	'font' => array(
		'bold' => true,
	)
);

$index = 1;
foreach($fields as $column => $name )
{

$objPHPExcel->setActiveSheetIndex(0)->setCellValueByColumnAndRow($column, $index, $name);

$letter = col2chr($column);
$objPHPExcel->getActiveSheet()->getStyle($letter.'1')->applyFromArray($styleArray);

if($name == "email")
	$objPHPExcel->getActiveSheet()->getColumnDimension( $letter )->setWidth(30);
elseif($name == "answer")
	$objPHPExcel->getActiveSheet()->getColumnDimension( $letter )->setWidth(8);
else
	$objPHPExcel->getActiveSheet()->getColumnDimension( $letter )->setWidth(20);

$phonecol = '';
if($name == 'phone')
	$phonecol = $letter;
}

	$sql = "SELECT * FROM ".$wpdb->prefix."rsvpmaker WHERE event=$eventid ORDER BY yesno DESC, last, first";
	$results = $wpdb->get_results($sql, ARRAY_A);
	$rows = sizeof($results);
	$maxcol = col2chr(sizeof($fields));
	$phonecells = $phonecol.'1:'.$phonecol.($rows+1);
	
$objPHPExcel->getActiveSheet()->getStyle($phonecells)->getNumberFormat()
->setFormatCode('###-###-####');

$bodyStyle = array(
	'borders' => array(
		'bottom' => array(
			'style' => PHPExcel_Style_Border::BORDER_THIN,
			'color' => array('argb' => '88888888'),
		)
	),
	'alignment' => array(
		'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
	)
);
	
	foreach($results as $row)
		{
		$index++;
		$objPHPExcel->getActiveSheet()->getStyle('A'.$index.':'.$maxcol.$index)->applyFromArray($bodyStyle);
		$row["yesno"] = ($row["yesno"]) ? "YES" : "NO";
		if($row["details"])
			{
			$details = unserialize($row["details"]);
			$row = array_merge($row,$details);
			}
		foreach($fields as $column => $name )
			{
				if(isset($row[$name]) )
					$objPHPExcel->setActiveSheetIndex(0)->setCellValueByColumnAndRow($column, $index, $row[$name]);
			}
			 //$worksheet->write($index, $column, $row[$name], $format_wrap);
		}

$objPHPExcel->getActiveSheet()->getStyle('A1:'.$maxcol.$rows)->getAlignment()->setWrapText(true)->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
$objPHPExcel->getActiveSheet()->getPageSetup()->setFitToWidth(1);
$objPHPExcel->getActiveSheet()->getPageSetup()->setFitToHeight(0);
$objPHPExcel->setActiveSheetIndex(0)->setCellValueByColumnAndRow(2, $index+3, "RSVPs for ".$title);
$objPHPExcel->getActiveSheet(0)->getHeaderFooter()->setOddHeader('&R RSVPs for  ' . $title . ' Page &P of &N');

header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="rsvp'.$eventid.'-'.date('Y-m-d-H-i').'.xls"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');

exit();
}

} // end rsvp_excel

add_action('admin_init','rsvp_excel');

?>