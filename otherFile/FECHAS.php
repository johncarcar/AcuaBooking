<?php
include 'PHPExcel_1.8/Classes/PHPExcel/IOFactory.php';
require 'PHPExcel_1.8/Classes/PHPExcel.php';

$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load('xls/Jet2/Book2.xlsx');

$objWorksheet  = $objPHPExcel->getActiveSheet();
// obtengo el valor de la celda
$fecha_excel = $objWorksheet->getCell('A50')->getValue();
// utilizo la función y obtengo el timestamp
$timestamp = PHPExcel_Shared_Date::ExcelToPHP($fecha_excel);
$fecha_php = date("Ymd",$timestamp);

echo $fecha_excel.'<br>';
echo $fecha_php.'<br>';



/*
//REALIZADO ANTERIORMENTE
//Creamos un objeto PHPExcel
$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load('xls/Jet2/Book2.xlsx');

// Indicamos que se pare en la hoja uno del libro
$objPHPExcel->setActiveSheetIndex(0);
$fecha= "=\"" ."\"";
$fecha=$objPHPExcel->setActiveSheetIndex()->getCell('A2')->getValue();
echo 'Valor de fecha antes de convertir = ';
echo $fecha.'<br>';

echo 'Valor de fecha despues de convertir ='. $fecha.'<br>';

*/



