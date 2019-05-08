
<?php
require 'PHPExcel_1.8/Classes/PHPExcel/IOFactory.php';
require 'PHPExcel_1.8/Classes/PHPExcel.php';


$nombreArchivo= 'xls/Jet2/Book1.xlsx';
$objPHPExcel= PHPExcel_IOFactory::load($nombreArchivo);
$objPHPExcel->setActiveSheetIndex(0);
$numRows=$objPHPExcel->setActiveSheetIndex(0)->getHighestRow();
for ($i=2; $i <= $numRows;$i++)

    {
        $fecha=$objPHPExcel->getActiveSheet()->getCell('A'.$i)->getCalculatedValue();
        $Bref=$objPHPExcel->getActiveSheet()->getCell('B'.$i)->getCalculatedValue();
        $nomRes=$objPHPExcel->getActiveSheet()->getCell('C'.$i)->getCalculatedValue();
        $personas=$objPHPExcel->getActiveSheet()->getCell('D'.$i)->getCalculatedValue();
        
        echo $fecha.','.$Bref.','.$nomRes.','.$personas.'<br>';   
            
    }    
    

