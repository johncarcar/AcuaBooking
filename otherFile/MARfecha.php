<?php
include 'C:/XAMPP/HTDOCS/BOOKINGS/PHPExcel_1.8/Classes/PHPExcel/IOFactory.php';
require 'C:/XAMPP/HTDOCS/BOOKINGS/PHPExcel_1.8/Classes/PHPExcel.php';

$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel5');
$objPHPExcel = $objReader->load('C:/XAMPP/HTDOCS/BOOKINGS/xls/Iberoservice/marismas2.xls');

$objWorksheet  = $objPHPExcel->getActiveSheet();
// obtengo el valor de la celda
$fecha_excel = $objWorksheet->getCell('C322')->getValue();
// utilizo la función y obtengo el timestamp
$timestamp = PHPExcel_Shared_Date::ExcelToPHP($fecha_excel);
$fecha_php = date("Ymd",$timestamp);

echo $fecha_excel.'<br>';
echo $fecha_php.'<br>';