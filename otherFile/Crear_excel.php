<?php
require 'PHPExcel_1.8/Classes/PHPExcel.php';

//MODULO QUE ESCRIBE EN UNA HOJA ACTIVA
$objPHPExcel= new PHPExcel();
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->setTitle('Hoja 1');
// Agregar en celda A1 valor string
$objPHPExcel->getActiveSheet()->setCellValue('A1','Prueba');
$objPHPExcel->getActiveSheet()->setCellValue('A2','aaaaa');
$objPHPExcel->getActiveSheet()->setCellValue('A3',TRUE);
$objPHPExcel->getActiveSheet()->setCellValue('A4','=20+50');
//IMPRIMIR
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="excel1.xls"');
header('Cache-Control:max-age=0');
$objPHPExcel= PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5');
$objPHPExcel->save('php://output');


