<!DOCTYPE html>
<html>
<head>
	<title>Leer Archivo Excel</title>
</head>
<body>
<h1>Leer Archivo Excel - IBS - MARISMAS - CABECERA</h1>
<?php
include 'PHPExcel_1.8/Classes/PHPExcel/IOFactory.php';
require 'PHPExcel_1.8/Classes/PHPExcel.php';
// Creamos un objeto PHPExcel
$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel5');
$objPHPExcel = $objReader->load('xls/Iberoservice/marismas.xls');
// Indicamos que se pare en la hoja uno del libro
$objPHPExcel->setActiveSheetIndex(0);
$allcell = $objPHPExcel->setActiveSheetIndex()->getHighestRow();

 $nueva=new PHPExcel_Worksheet($objPHPExcel,'sheet2');
 $objPHPExcel->addSheet($nueva,1);
 
 //inicia nuevas variables
 //$letra='A';   
 $j=1;
 $Nalba=$_GET['numero'];
 $valold='';
for ($i=1;$i<=$allcell;$i++){
    $varM= $objPHPExcel->setActiveSheetIndex(0)->getCell('M'.$i);
    $varV= $objPHPExcel->setActiveSheetIndex(0)->getCell('V'.$i);
    $varC= $objPHPExcel->setActiveSheetIndex(0)->getCell('A'.$i); //VARIABLE QUE COJE LA FECHA

    if ($varM<>'' and $varV<>'' and $varM<>$valold){
        //FECHA
        $objWorksheet  = $objPHPExcel->getActiveSheet(0);
        // obtengo el valor de la celda
        $fecha_excel = $objWorksheet->getCell('A'.$i)->getValue();
        // utilizo la función y obtengo el timestamp
        $timestamp = PHPExcel_Shared_Date::ExcelToPHP($fecha_excel);
        $fecha_php = date("dmY",$timestamp);
        //// echo$fecha_excel.' - '; //muestra la fecha valor excel
      
        
        //escribe en el nuevo libro
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('A'.$j,'V');
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('B'.$j,'A');
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,'AV');
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('E'.$j,$fecha_php);
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('F'.$j,'4300000016');
        $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,$varM);
        // echo'V - A - ALB -'.$Nalba.' - '.$fecha_php.'- 4300000016 '.$varM.'<br>';
        //asigna valor a $valold y se compare con el nuevo valor
        $valold=$objPHPExcel->setActiveSheetIndex(0)->getCell('M'.$i);
        $Nalba++;          
        $j++;
    }
}
//Inserta linea con datos de cabecera;
$objPHPExcel->setActiveSheetIndexByName('sheet2');
$objPHPExcel->getActiveSheet()->insertNewRowBefore(1);
$objPHPExcel->getActiveSheet()->setCellValue('A1','CLASE');
$objPHPExcel->getActiveSheet()->setCellValue('B1','TIPO');
$objPHPExcel->getActiveSheet()->setCellValue('C1','SERIE');
$objPHPExcel->getActiveSheet()->setCellValue('D1','NALBA');
$objPHPExcel->getActiveSheet()->setCellValue('E1','FECHA');
$objPHPExcel->getActiveSheet()->setCellValue('F1','CLIENTE');
$objPHPExcel->getActiveSheet()->setCellValue('G1','REFERENCIA');

$objPHPExcel->removeSheetByIndex(0);
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("xls/Iberoservice/CABmarismas.xlsx");
//llamar al archivo que crea las lineas
$Nalba--;
echo' El último albaran creado es el: '.$Nalba.'<br>';


require_once 'MARlineas.php';

?>

<p>Archivo Exportado<p> 
<h2><a href="index.php">Volver a Index</a></h2>      
</body>
                 
</html>
