<?php

include 'C:/XAMPP/HTDOCS/BOOKINGS/PHPExcel_1.8/Classes/PHPExcel/IOFactory.php';
require 'C:/XAMPP/HTDOCS/BOOKINGS/PHPExcel_1.8/Classes/PHPExcel.php';
// Creamos un objeto PHPExcel
$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel5');
$objPHPExcel = $objReader->load('C:/XAMPP/HTDOCS/BOOKINGS/xls/Iberoservice/marismas2.xls');
// Indicamos que se pare en la hoja uno del libro
$objPHPExcel->setActiveSheetIndex(0);
$allcell = $objPHPExcel->setActiveSheetIndex()->getHighestRow();

 $nueva=new PHPExcel_Worksheet($objPHPExcel,'sheet2');
 $objPHPExcel->addSheet($nueva,1);
 
 //inicia nuevas variables
 //$letra='A';   
 $j=1;
 $Nalba=500;
 $valold='';
for ($i=1;$i<=$allcell;$i++){
    $varM= $objPHPExcel->setActiveSheetIndex(0)->getCell('M'.$i);
    $varV= $objPHPExcel->setActiveSheetIndex(0)->getCell('V'.$i);
    $varC= $objPHPExcel->setActiveSheetIndex(0)->getCell('C'.$i);

    if ($varM<>'' and $varV<>'' and $varM<>$valold){
        echo $varM.' - ';
        //FECHA
        $objWorksheet  = $objPHPExcel->getActiveSheet(0);
        // obtengo el valor de la celda
        $fecha_excel = $objWorksheet->getCell('C'.$i)->getValue();
        // utilizo la función y obtengo el timestamp
        $timestamp = PHPExcel_Shared_Date::ExcelToPHP($fecha_excel);
        $fecha_php = date("Ymd",$timestamp);
        //echo $fecha_excel.' - '; //muestra la fecha valor excel
        echo $fecha_php.' - '; //muestra la fecha convertida a valor php      
        
        //escribe en el nuevo libro
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('A'.$j,'V');
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('B'.$j,'A');
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,'ALB');
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,$varM);
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('E'.$j,$fecha_php);
        //asigna valor a $valold y se compare con el nuevo valor
        $valold=$objPHPExcel->setActiveSheetIndex(0)->getCell('M'.$i);
        $Nalba++;          
        $j++;
    }
}

$objPHPExcel->removeSheetByIndex(0);
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("C:/XAMPP/HTDOCS/BOOKINGS/xls/Iberoservice/CABmarismas.xlsx");
