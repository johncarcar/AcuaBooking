<!DOCTYPE html>
<html>
<head>
	<title>Leer Archivo Excel</title>
</head>
<body>
<h1>RESERVAS - Leer lineas de jet2 Archivo Excel</h1>
<p>Archivo Exportado<p>
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


//Borramos la línea de encabezados
$objPHPExcel->setActiveSheetIndex()->removeRow(1);

//INSERTAMOS COLUMNAS DE LA CABECERA.
//CLASE = V - TIPO = A - SERIE =ALB - ALBANUMERO = ALB - CLIENTE = 4300000025

for ($i=1;$i<=10;$i++){ //Creamos las columnas que se ván a rellenar.
    $objPHPExcel->setActiveSheetIndex()->insertNewColumnBefore('A');
    }
$numRows=$objPHPExcel->setActiveSheetIndex(0)->getHighestRow('O');


//Rellenar campos que por cada 4 insertamos 

    $Nalba='100';
    $num=1;
    $linea=1;
    while ($linea<=$numRows){
            //Pocisión de la Linea a importar.
            
            //Inserta la Primera Línea del Albaran que contiene el nombre del cliente.
            $objPHPExcel->getActiveSheet()->setCellValue('A'.$num,'V');
            $objPHPExcel->getActiveSheet()->setCellValue('B'.$num,'A');
            $objPHPExcel->getActiveSheet()->setCellValue('C'.$num,'ALB');
            $objPHPExcel->getActiveSheet()->setCellValue('D'.$num,$Nalba);
            $objPHPExcel->getActiveSheet()->setCellValue('E'.$num,'DESC');
            //inserto el valor de la celda a grabar en una variable.
            $Desc =$objPHPExcel->getActiveSheet()->getCell('M'.$linea);
            $objPHPExcel->getActiveSheet()->setCellValue('F'.$num,$Desc);
            $objPHPExcel->getActiveSheet()->setCellValue('G'.$num,'');
            $objPHPExcel->getActiveSheet()->setCellValue('H'.$num,'');
            $num++;
            //Inserta la Primera Línea del Albaran que contiene los Acompañantes.
            $objPHPExcel->getActiveSheet()->setCellValue('A'.$num,'V');
            $objPHPExcel->getActiveSheet()->setCellValue('B'.$num,'A');
            $objPHPExcel->getActiveSheet()->setCellValue('C'.$num,'ALB');
            $objPHPExcel->getActiveSheet()->setCellValue('D'.$num,$Nalba);
            $objPHPExcel->getActiveSheet()->setCellValue('E'.$num,'ACOMP');
            //inserto el valor de la celda a grabar en una variable.
            $Acomp =$objPHPExcel->getActiveSheet()->getCell('N'.$linea);
            $objPHPExcel->getActiveSheet()->setCellValue('F'.$num,$Acomp);
            $objPHPExcel->getActiveSheet()->setCellValue('G'.$num,'');
            $objPHPExcel->getActiveSheet()->setCellValue('H'.$num,'');
            $num++;
            //Inserta la Primera Línea del Albaran que contiene Cantidad ADULTOS
            $objPHPExcel->getActiveSheet()->setCellValue('A'.$num,'V');
            $objPHPExcel->getActiveSheet()->setCellValue('B'.$num,'A');
            $objPHPExcel->getActiveSheet()->setCellValue('C'.$num,'ALB');
            $objPHPExcel->getActiveSheet()->setCellValue('D'.$num,$Nalba);
            $objPHPExcel->getActiveSheet()->setCellValue('E'.$num,'7770200009');
            //inserto el valor de la celda a grabar en una variable.
            $Adult =$objPHPExcel->getActiveSheet()->getCell('Q'.$linea);
            $objPHPExcel->getActiveSheet()->setCellValue('F'.$num,'JET2 ADULTO DEST');
            $objPHPExcel->getActiveSheet()->setCellValue('G'.$num,$Adult);
            $objPHPExcel->getActiveSheet()->setCellValue('H'.$num,'18,75');       
            $num++;
            //Inserta la Primera Línea del Albaran que contiene Cantidad NIÑOS
            $objPHPExcel->getActiveSheet()->setCellValue('A'.$num,'V');
            $objPHPExcel->getActiveSheet()->setCellValue('B'.$num,'A');
            $objPHPExcel->getActiveSheet()->setCellValue('C'.$num,'ALB');
            $objPHPExcel->getActiveSheet()->setCellValue('D'.$num,$Nalba);
            $objPHPExcel->getActiveSheet()->setCellValue('E'.$num,'7770200008');
            //inserto el valor de la celda a grabar en una variable.
            $Nino =$objPHPExcel->getActiveSheet()->getCell('R'.$linea);
            $objPHPExcel->getActiveSheet()->setCellValue('F'.$num,'JET2 NIÑO DEST');
            $objPHPExcel->getActiveSheet()->setCellValue('G'.$num,$Nino);
            $objPHPExcel->getActiveSheet()->setCellValue('H'.$num,'14,25');       
            $num++;
            //Sumamos una línea para que siga leyendo el archivo
            $linea++;
            $Nalba++;
    }    

//BORRAMOS LAS COLUMNAS QUE SOBRAN..
$objPHPExcel->setActiveSheetIndex()->removeColumn('K','L','M','N','O','P','Q','R','T','U','V','W','X','Y','Z'); //Borramos las lineas restantes

//Inserta linea con datos de cabecera;
$objPHPExcel->getActiveSheet()->insertNewRowBefore(1);
$objPHPExcel->getActiveSheet()->setCellValue('A1','CLASE');
$objPHPExcel->getActiveSheet()->setCellValue('B1','TIPO');
$objPHPExcel->getActiveSheet()->setCellValue('C1','SERIE');
$objPHPExcel->getActiveSheet()->setCellValue('D1','NALBA');
$objPHPExcel->getActiveSheet()->setCellValue('E1','CODIGO');
$objPHPExcel->getActiveSheet()->setCellValue('F1','DESCR');
$objPHPExcel->getActiveSheet()->setCellValue('G1','CANTIDAD');
$objPHPExcel->getActiveSheet()->setCellValue('H1','IMPORTE');

//Guardamos el documento.
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("x:/TTOO/xls/JET2/Lineas_JET2.xlsx"); 
?>


</body>
                 
</html>
