<?php
include 'PHPExcel_1.8/Classes/PHPExcel/IOFactory.php';
require 'PHPExcel_1.8/Classes/PHPExcel.php';
// Creamos un objeto PHPExcel
$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load('xls/Jet2/Book2.xlsx');
// Indicamos que se pare en la hoja uno del libro
$objPHPExcel->setActiveSheetIndex(0);

//BORRAMOS LAS COLUMNAS..
$objPHPExcel->setActiveSheetIndex()->removeColumn('E','F','G','H','I','J','K');
//Borramos la línea de encabezados
$objPHPExcel->setActiveSheetIndex()->removeRow(1);

//INSERTAMOS COLUMNAS DE LA CABECERA.
//CLASE = V - TIPO = A - SERIE =ALB - ALBANUMERO = ALB - CLIENTE = 4300000025

for ($i=1;$i<=5;$i++){ //Creamos las columnas que se ván a rellenar.
    $objPHPExcel->setActiveSheetIndex()->insertNewColumnBefore('A');
    }
$numRows=$objPHPExcel->setActiveSheetIndex(0)->getHighestRow();
$Nalba='100';
for ($j=1;$j<=$numRows;$j++){    
    $objPHPExcel->getActiveSheet()->setCellValue('A'.$j,'V');
    $objPHPExcel->getActiveSheet()->setCellValue('B'.$j,'A');
    $objPHPExcel->getActiveSheet()->setCellValue('C'.$j,'ALB');
    $objPHPExcel->getActiveSheet()->setCellValue('D'.$j,$Nalba);
    $objPHPExcel->getActiveSheet()->setCellValue('E'.$j,'4300000025');
    $Nalba++;
    
}    
//Guardamos el documento.
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("Cabecera.xlsx");


