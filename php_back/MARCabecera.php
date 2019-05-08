<!DOCTYPE html>
<html>
<head>
	<title>Leer Archivo Excel</title>
</head>
<body>
<h1>Leer Archivo Excel - IBERO - MARISMAS - CABECERA</h1>
<?php
include 'PHPExcel_1.8/Classes/PHPExcel/IOFactory.php';
require 'PHPExcel_1.8/Classes/PHPExcel.php';
// Creamos un objeto PHPExcel
$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel5');
$objPHPExcel = $objReader->load('xls/Iberoservice/marismas2.xls');
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
        //FECHA
        $objWorksheet  = $objPHPExcel->getActiveSheet(0);
        // obtengo el valor de la celda
        $fecha_excel = $objWorksheet->getCell('C'.$i)->getValue();
        // utilizo la función y obtengo el timestamp
        $timestamp = PHPExcel_Shared_Date::ExcelToPHP($fecha_excel);
        $fecha_php = date("dmY",$timestamp);
        //echo $fecha_excel.' - '; //muestra la fecha valor excel
      
        
        //escribe en el nuevo libro
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('A'.$j,'V');
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('B'.$j,'A');
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,'ALB');
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('E'.$j,$fecha_php);
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('F'.$j,'4300000016');
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,$varM);
        echo 'V - A - ALB -'.$Nalba.' - '.$fecha_php.'- 4300000016 '.$varM.'<br>';
        //asigna valor a $valold y se compare con el nuevo valor
        $valold=$objPHPExcel->setActiveSheetIndex(0)->getCell('M'.$i);
        $Nalba++;          
        $j++;
    }
}

$objPHPExcel->removeSheetByIndex(0);
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("xls/Iberoservice/CABmarismas.xlsx");

?>
<p>Archivo Exportado<p> 
</body>
                 
</html>
