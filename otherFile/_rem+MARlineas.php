<!DOCTYPE html>
<html>
<head>
	<title>Leer Archivo Excel</title>
</head>
<body>
<h1>RESERVAS - Leer lineas de IBEROSERVICE - MARISMAS -  Archivo Excel</h1>

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

// variables a utilizar con sus valores:
$codArt='';
$codProvA='';
$codProvN='';
$descArtA='';
$descArtN='';


$Nalba=500;
$j=1; 

for ($i=1;$i<=$allcell;$i++){
    $varM= $objPHPExcel->setActiveSheetIndex(0)->getCell('M'.$i);
    $varV= $objPHPExcel->setActiveSheetIndex(0)->getCell('V'.$i);
    $varAN= $objPHPExcel->setActiveSheetIndex(0)->getCell('AN'.$i);
    //$valold=$objPHPExcel->setActiveSheetIndex(0)->getCell('M'.$i);
    
        if ($varM<>'' and $varAN<>''){   
            //escribe en el nuevo libro
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,$varM);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,$varV);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('F'.$j,$varAN);
            //asigna valor a $valold y se compare con el nuevo valor
            //$valold=$objPHPExcel->setActiveSheetIndex(0)->getCell('M'.$i);
             echo $varM.' Caso1 - ';
             echo $Nalba.' - ';
             $Nalba++;
        }
        if ($varM=='' and $varAN<>''){
            $Nalba--;
            //escribe en el nuevo libro
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,$varM);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,$varV);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('F'.$j,$varAN);
            //asigna valor a $valold y se compare con el nuevo valor
            //$valold=$objPHPExcel->setActiveSheetIndex(0)->getCell('M'.$i);
            echo $varM.' Caso2 - ';
            echo $Nalba.' - ';
            $Nalba++;
            
        }
        if (($varM=='' and $varAN=='')){
            
            echo $varM.' Caso3 - ';
            echo $Nalba.' - ';
            $j--;
        }  
        $j++;
        
}

$objPHPExcel->removeSheetByIndex(0);
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("C:/XAMPP/HTDOCS/BOOKINGS/xls/Iberoservice/LINmarismas.xlsx");

?>

<p>Archivo Exportado<p>
</body>
                 
</html>
