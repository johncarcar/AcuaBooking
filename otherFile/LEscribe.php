
<?php

include 'PHPExcel_1.8/Classes/PHPExcel/IOFactory.php';
require 'PHPExcel_1.8/Classes/PHPExcel.php';

//LECTURA
$nombreArchivo= 'xls/Jet2/Book1.xlsx';
$objPHPExcel= PHPExcel_IOFactory::load($nombreArchivo);
$objPHPExcel->setActiveSheetIndex(0);
$numRows=$objPHPExcel->setActiveSheetIndex(0)->getHighestRow();

for ($i=2; $i <= 3;$i++)

    {
        $fecha=$objPHPExcel->getActiveSheet()->getCell('A'.$i)->getCalculatedValue();
        $Bref=$objPHPExcel->getActiveSheet()->getCell('B'.$i)->getCalculatedValue();
        $nomRes=$objPHPExcel->getActiveSheet()->getCell('C'.$i)->getCalculatedValue();
        $personas=$objPHPExcel->getActiveSheet()->getCell('D'.$i)->getCalculatedValue();
        
        echo $fecha.','.$Bref.','.$nomRes.','.$personas.'<br>';   
        
        
    } 
    
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("Archivo_salida.xlsx");

/*    
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="excel1.xls"');
header('Cache-Control:max-age=0');
$objPHPExcel= PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5');
$objPHPExcel->save('php://output');
*/