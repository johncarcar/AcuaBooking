<!DOCTYPE html>
<html>
<head>
	<title>Leer Archivo Excel</title>
</head>
<body>
<h1>Leer Archivo Excel</h1>
<?php
include 'PHPExcel_1.8/Classes/PHPExcel/IOFactory.php';
require 'PHPExcel_1.8/Classes/PHPExcel.php';
// Creamos un objeto PHPExcel
$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load('x:/TTOO/xls/Jet2/Book2.xlsx');
// Indicamos que se pare en la hoja uno del libro
$objPHPExcel->setActiveSheetIndex(0);

//BORRAMOS LAS COLUMNAS..
$objPHPExcel->setActiveSheetIndex()->removeColumn('C','D','E','F','G','H','I','J','K');
//Borramos la línea de encabezados
$objPHPExcel->setActiveSheetIndex()->removeRow(1);

//INSERTAMOS COLUMNAS DE LA CABECERA.
//CLASE = V - TIPO = A - SERIE =ALB - ALBANUMERO = ALB - CLIENTE = 4300000025

for ($i=1;$i<=6;$i++){ //Creamos las columnas que se ván a rellenar.
    $objPHPExcel->setActiveSheetIndex()->insertNewColumnBefore('A');
    }
$numRows=$objPHPExcel->setActiveSheetIndex(0)->getHighestRow();
$Nalba='100';
for ($j=1;$j<=$numRows;$j++){

    $objPHPExcel->getActiveSheet()->setCellValue('A'.$j,'V');
    $objPHPExcel->getActiveSheet()->setCellValue('B'.$j,'A');
    $objPHPExcel->getActiveSheet()->setCellValue('C'.$j,'ALB');
    
    //Cambiamos el valor recibido de la fecha, luego, inserta la fecha,
    $objWorksheet  = $objPHPExcel->getActiveSheet();
     
    
    $fecha_excel = $objWorksheet->getCell('G'.$j)->getValue();
    $timestamp = PHPExcel_Shared_Date::ExcelToPHP($fecha_excel);
    $fecha_php = date("Ymd",$timestamp);
    $objPHPExcel->getActiveSheet()->setCellValue('D'.$j,$fecha_php);
    //$objPHPExcel->setActiveSheetIndex()->removeColumn('G');
    //Sigue insertando datos.
    $objPHPExcel->getActiveSheet()->setCellValue('E'.$j,$Nalba);
    $objPHPExcel->getActiveSheet()->setCellValue('F'.$j,'4300000025');
    $Nalba++;
 
}
//elimina columna sobrante
$objPHPExcel->setActiveSheetIndex()->removeColumn('G');

//Inserta linea con datos de cabecera;
$objPHPExcel->getActiveSheet()->insertNewRowBefore(1);
$objPHPExcel->getActiveSheet()->setCellValue('A1','CLASE');
$objPHPExcel->getActiveSheet()->setCellValue('B1','TIPO');
$objPHPExcel->getActiveSheet()->setCellValue('C1','SERIE');
$objPHPExcel->getActiveSheet()->setCellValue('D1','FECHA');
$objPHPExcel->getActiveSheet()->setCellValue('E1','NALBA');
$objPHPExcel->getActiveSheet()->setCellValue('F1','CLIENTE');
$objPHPExcel->getActiveSheet()->setCellValue('G1','REFERENCIA');
//Guardamos el documento.
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("x:/TTOO/xls/JET2/cabecera_JET2.xlsx"); 
?>

<p>Archivo Exportado<p>
</body>
                 
</html>
