<?php
//  Required - $_REQUEST[data] (json_encoded - 2 dimensional array)
if(!$_REQUEST['data'])
    die('false');
error_reporting(0);
require_once(dirname(__FILE__).'/../shared_data/phpExcel/PHPExcel.php');
$in_data = $_REQUEST['data'];
try{
    $in_data = json_decode(trim($in_data));
}catch(Exception $ex){
    die($ex);
}

$_REQUEST['title'] = ($_REQUEST['title']) ? $_REQUEST['title'] : "Excel";

$excel = new PHPExcel();
$excel->getProperties()->setCreator("derp@gmail.com")
                        ->setLastModifiedBy("derp@gmail.com")
                        ->setTitle("")
                        ->setSubject("")
                        ->setDescription("")
                        ->setKeywords("")
                        ->setCategory("");
$excel->getDefaultStyle()->getFont()->setName('Calibri');
$excel->getDefaultStyle()->getFont()->setSize(11);
$excel->setActiveSheetIndex(0)->setTitle($_REQUEST['title']);
$headers_style = array(
            'borders' => array(
                    'outline' => array(
                            'style' => PHPExcel_Style_Border::BORDER_THIN,
                            'color' => array('argb' => '22251d'),
                    ),
                    'bottom' => array(
                            'style' => PHPExcel_Style_Border::BORDER_MEDIUM,
                            'color' => array('argb' => '001670'),
                    ),
            ),
            'fill' 	=> array(
                    'type'		=> PHPExcel_Style_Fill::FILL_SOLID,
                    'color'		=> array('argb' => 'c4d79b')
            ),
            'font'=>array(
				'name'      =>  'Calibri',
				'size'      =>  12,
				'bold'      => true
			),
    );
$inner_borders = array(
        'borders' => array(
                'allborders' => array(
                        'style' => PHPExcel_Style_Border::BORDER_THIN,
                        'color' => array('argb' => '22251d'),
                ),
        ),
);
$first = array_slice($in_data,0,1);
$first = $first[0];
$alphabet = $final_alphabet = range("A","Z");
$loopCnt = 0;
$spreadsheet_row = 1;
while(count($first) > count($final_alphabet)){
    $working_letter = array_slice($alphabet,$loopCnt,1);
    $working_letter = $working_letter[0];
    foreach($alphabet as $letter){
        array_push($final_alphabet,$alphabet[$loopCnt].$letter);
    }
}
foreach(array_keys((array)$first) as $num=>$key){
    $excel->getActiveSheet()->setCellValue($final_alphabet[$num].$spreadsheet_row, $key);
    $final_letter = $final_alphabet[$num];
}
foreach($in_data as $data_row){
    $spreadsheet_row++;
    $num = 0;
    foreach($data_row as $data_col){
        $excel->getActiveSheet()->setCellValue($final_alphabet[$num].$spreadsheet_row, $data_col);
        $num++;
    }
}
$excel->getActiveSheet()->getStyle("A1:".$final_letter."1")->applyFromArray($headers_style);
$excel->getActiveSheet()->getStyle("A1:".$final_letter.($spreadsheet_row))->applyFromArray($inner_borders);
$objWriter  = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
ob_start();
$objWriter->save('php://output');
header('Content-Type: application/vnd.openXMLformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="'.$_REQUEST['title'].'.xlsx"');
header('Cache-Control: max-age=0');
echo ob_get_clean();
?>