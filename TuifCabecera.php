<!-- ULTIMO CAMBIO 24/04/2019 - V3 -->
<!DOCTYPE html>
<html>
<head>
	<title>Leer Archivo Excel</title>
</head>
<body>
<h1>Leer Archivo Excel - TUI FRANCE - CABECERA</h1>
<?php
include 'PHPExcel_1.8/Classes/PHPExcel/IOFactory.php';
require 'PHPExcel_1.8/Classes/PHPExcel.php';
// Creamos un objeto PHPExcel
$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load('xls/Tfrance/TuiF.xlsx');
// Indicamos que se pare en la hoja uno del libro
$objPHPExcel->setActiveSheetIndex(0);
$allcell = $objPHPExcel->setActiveSheetIndex()->getHighestRow();

// Indicamos que se pare en la hoja uno del libro
$objPHPExcel->setActiveSheetIndex(0);
$allcell = $objPHPExcel->setActiveSheetIndex()->getHighestRow();

//Separa todas las celdas que se encuentran juntas (UNMERGECELL)
$celdas=$objPHPExcel->setActiveSheetIndex(0)->getMergeCells();
foreach($celdas as $value){
    $objPHPExcel->setActiveSheetIndex(0)->unmergeCells($value);
} 

// LO siguiente es calcular la cantidad de filas y columnas que tiene la hoja rellenas.
 $numRows=$objPHPExcel ->setActiveSheetIndex(0) ->getHighestRow();
 
//Se agrega una nueva hoja y se le asigna el index 1   
 $nueva=new PHPExcel_Worksheet($objPHPExcel,'sheet');
 $objPHPExcel->addSheet($nueva,1);

 //Inicio de la comprovación de datos.
 $Nalba= $_GET['numero'];
 $i=1;
 for ($j = 10; $j <= $numRows; $j++) {//SE EMPIESA A COGER DESDE LA LINEA 10
    $fecha_excel = $objPHPExcel->setActiveSheetIndex(0)->getCell('G' . $j)->getValue(); //OBTIENE VALOR DE CELDA FECHA
    $timestamp = PHPExcel_Shared_Date::ExcelToPHP($fecha_excel); // HACE TIMESTAMP DE LA FECHA
    $fecha_php = date("dmY", $timestamp);   //FECHA ESCOGIDA   // ASIGNA ESE TIME STAMP A UNA NUEVA VARIABLE DANDO FROMATO
    $Bref = $objPHPExcel->setActiveSheetIndex(0)->getCell('J' . $j)->getValue(); //BOOKING REFERENCE
    $edad = $objPHPExcel->setActiveSheetIndex(0)->getCell('S' . $j)->getValue(); //VARIABLE EDAD
    //INTRODUCCIÓN DE DATOS
    if ($Bref <> 'Subtotal') {//SE DESCARTAN LAS LINEAS QUE TIENEN SUBTOTAL EN BOOKING REFERENCE
        if ($fecha_excel <> '' & $Bref <> '') {
            $fecha_old = $fecha_php; $Nalba_old = $Nalba; $Bref_old = $Bref;
            echo $Nalba . ' - ' . $fecha_php . ' - ' . $Bref . ' - ' . $edad . '<br>';
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('A' . $i, 'V');
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('B' . $i, 'A');            
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('C' . $i, 'AV');
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('D' . $i, $Nalba);
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('E' . $i, $fecha_php); //LLENADOD DE LAS FECHAS EN EL SISTEMA
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('F' . $i, '4300000008');            
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('G' . $i, $Bref); // LLENANDO BOOKING REF EN EL SISTEMA  
            $Nalba++; 
            $i++;
        }
        if ($fecha_excel == '' & $Bref <> '') {
            $Bref_old = $Bref; $Nalba_old= $Nalba; //Variable temporal que guarda el valor elegido.
            echo $Nalba . ' - ' . $fecha_old . ' - ' . $Bref . ' - ' . $edad . '<br>';
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('A' . $i, 'V');
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('B' . $i, 'A');
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('C' . $i, 'AV');
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('D' . $i, $Nalba);
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('E' . $i, $fecha_old); //LLENADOD DE LAS FECHAS EN EL SISTEMA
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('F' . $i, '4300000008');
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('G' . $i, $Bref); // LLENANDO BOOKING REF EN EL SISTEMA  
            $i++;
            $Nalba++;
        }
        /* SE DESCARTÓ PORQUE SOLO SE TRAE UNA LÍNEA POR ALBARÁN
        if ($fecha_excel == '' & $Bref == '' & $edad <> '') {
            echo $Nalba_old . ' - ' . $fecha_old . ' - ' . $Bref_old . ' - ' . $edad . '<br>';
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('A' . $i, 'V');
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('B' . $i, 'A');
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('C' . $i, 'AV');
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('D' . $i, $Nalba_old);
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('E' . $i, $fecha_old);
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('F' . $i, '4300000008');
            $objPHPExcel->setActiveSheetIndex(1)->setCellValue('G' . $i, $Bref_old); 
            $i++;
        }*/
    }
}
//Inserta linea con datos de cabecera;
$objPHPExcel->setActiveSheetIndexByName('sheet');
$objPHPExcel->getActiveSheet()->insertNewRowBefore(1);
$objPHPExcel->getActiveSheet()->setCellValue('A1','CLASE');
$objPHPExcel->getActiveSheet()->setCellValue('B1','TIPO');
$objPHPExcel->getActiveSheet()->setCellValue('C1','SERIE');
$objPHPExcel->getActiveSheet()->setCellValue('D1','NALBA');
$objPHPExcel->getActiveSheet()->setCellValue('E1','FECHA');
$objPHPExcel->getActiveSheet()->setCellValue('F1','CLIENTE');
$objPHPExcel->getActiveSheet()->setCellValue('G1','REFERENCIA');


$objPHPExcel->removeSheetByIndex(0);
$objPHPExcel->removeSheetByIndex(1);

//Guardar las modificaciones 
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('xls/Tfrance/CabTuiF.xlsx');

//llamar al archivo que crea las lineas
require_once 'TuifLineas.php';
?>

<h2><a href="index_1.php">Volver a Index</a></h2>
</body>   
</html>
