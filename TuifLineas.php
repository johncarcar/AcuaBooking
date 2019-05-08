<!-- ULTIMO CAMBIO 24/04/2019 - V3 -->
<!DOCTYPE html>
<html>
<head>
	<title>Leer Archivo Excel</title>
</head>
<body>
<h1>Leer Archivo Excel - TUI FRANCE - LINEAS </h1>
<?php
//include 'PHPExcel_1.8/Classes/PHPExcel/IOFactory.php'; // SE DEBEN DE DESACTIVAR PARA QUE SEAN LLAMADOS DESDE CABECERA
//require 'PHPExcel_1.8/Classes/PHPExcel.php';          // SE DEBEN DE DESACTIVAR PARA QUE SEAN LLAMADOS DESDE CABECERA
// Creamos un objeto PHPExcel
$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load('xls/Tfrance/TuiF.xlsx');
// Indicamos que se pare en la hoja uno del libro
$objPHPExcel->setActiveSheetIndex(0);
$allcell = $objPHPExcel->setActiveSheetIndex(0)->getHighestRow();

// Indicamos que se pare en la hoja uno del libro
$objPHPExcel->setActiveSheetIndex(0);
$allcell = $objPHPExcel->setActiveSheetIndex(0)->getHighestRow();

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
 
 //inicia nuevas variables. Variables a utilizar con sus valores:
    $codPrecNS='24';
    $codPrecA='49';
    $codArtA='7770200056';
    $codArtNS='7770200055';
    $descArtA='TUI FR ADULTO BAJA 19';
    $descArtNS='TUI FR NIÑO/SEN BAJA 19';
 
 //Variables que se van a colocar en los documentos
    $precio='0';
    $art='';
    $desc='';

    
 //Inicio de la comprovación de datos.
 $Nalba=$_GET['numero'];
 $i=1;
 for ($j = 10; $j <= $numRows; $j++) { //SE EMPIESA A COGER DESDE LA LINEA 10
    $fecha_excel = $objPHPExcel->setActiveSheetIndex(0)->getCell('G' . $j)->getValue(); //OBTIENE VALOR DE CELDA FECHA
    $timestamp = PHPExcel_Shared_Date::ExcelToPHP($fecha_excel); // HACE TIMESTAMP DE LA FECHA
    $fecha_php = date("dmY", $timestamp);   //FECHA ESCOGIDA   // ASIGNA ESE TIME STAMP A UNA NUEVA VARIABLE DANDO FROMATO
    $Bref = $objPHPExcel->setActiveSheetIndex(0)->getCell('J' . $j)->getValue(); //BOOKING REFERENCE
    $edad = $objPHPExcel->setActiveSheetIndex(0)->getCell('S' . $j)->getValue(); //VARIABLE EDAD
    $varNomb=$objPHPExcel->setActiveSheetIndex(0)->getCell('Q' . $j)->getValue(); //VARIABLE EDAD
    //INTRODUCCIÓN DE DATOS
    
    if ($edad>64 OR ($edad >3 && $edad<12)){
        $art=$codArtNS;
        $desc=$descArtNS;
        $precio=$codPrecNS;
    }
    if ($edad>11 && $edad<65){
        $art=$codArtA;
        $desc=$descArtA;
        $precio=$codPrecA;
    }
    
    if ($edad>3 or $edad==''){ //DESCARTE DE BEBES Y VACIOS
        if ($Bref <> 'Subtotal') {//SE DESCARTAN LAS LINEAS QUE TIENEN SUBTOTAL EN BOOKING REFERENCE
            if ($fecha_excel <> '' & $Bref <> '') {
                $Bref_old = $Bref; $Nalba_old= $Nalba; $fecha_old=$fecha_php; //Variable temporal que guarda el valor elegido.
                echo $Nalba . ' - ' . $Bref_old. ' - ' .$edad .' - '.$art.' - '.$desc.'-'.$varNomb.' - '.$precio.' - '.$i.'<br>';
                
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('A' . $i, 'V');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('B' . $i, 'A');            
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('C' . $i, 'AV');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('D' . $i, $Nalba);
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('E' . $i, $art); //CÓDIGO DEL ARTÍCULO
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('F' . $i, $desc);//DESCRIPCIÓN DEL ARTÍCULO
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('G' . $i, '1');//CANTIDAD DEL ARTÍCULO
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('H' . $i, $precio);//PRECIO DEL ARTÍCULO
                $i++;
                
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('A' . $i, 'V');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('B' . $i, 'A');            
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('C' . $i, 'AV');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('D' . $i, $Nalba);
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('E' . $i,'DESC');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('F' . $i,$varNomb);
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('G' . $i,'');           
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('H' . $i,'');
                $Nalba++; 
                $i++;                
                   
        }
            if ($fecha_excel == '' & $Bref <> '') {
                $Bref_old = $Bref; $Nalba_old= $Nalba; $fecha_old=$fecha_php;//Variable temporal que guarda el valor elegido.
                echo $Nalba. ' - '. $Bref. ' - ' .$edad .' - '.$art.' - '.$desc.'-'.$varNomb.' - '.$precio.' - '.$i.'<br>';
                
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('A' . $i, 'V');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('B' . $i, 'A');            
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('C' . $i, 'AV');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('D' . $i, $Nalba);
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('E' . $i, $art); //CÓDIGO DEL ARTÍCULO
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('F' . $i, $desc);//DESCRIPCIÓN DEL ARTÍCULO
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('G' . $i, '1');//CANTIDAD DEL ARTÍCULO
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('H' . $i, $precio);//PRECIO DEL ARTÍCULO
                $i++;
                
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('A'.$i, 'V');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('B'.$i, 'A');            
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('C'.$i, 'AV');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('D'.$i, $Nalba);
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('E'.$i,'DESC');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('F'.$i,$varNomb);
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('G'.$i,'');           
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('H'.$i,'');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('I'.$i,'');
                $Nalba++; 
                $i++;                
            }
            if ($fecha_excel == '' & $Bref == '' & $edad <> '') {
                
                echo $Nalba_old . ' - ' . $Bref_old. ' - ' .$edad .' - '.$art.' - '.$desc.'-'.$varNomb.' - '.$precio.' - '.$i.'<br>';
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('A' . $i, 'V');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('B' . $i, 'A');            
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('C' . $i, 'AV');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('D' . $i, $Nalba_old);
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('E' . $i,$art); //CÓDIGO DEL ARTÍCULO
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('F' . $i, $desc);//DESCRIPCIÓN DEL ARTÍCULO
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('G' . $i, '1');//CANTIDAD DEL ARTÍCULO
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('H' . $i, $precio);//PRECIO DEL ARTÍCULO
                $i++;
                
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('A' . $i, 'V');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('B' . $i, 'A');            
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('C' . $i, 'AV');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('D' . $i, $Nalba_old);
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('E' . $i,'DESC');
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('F' . $i,$varNomb);
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('G' . $i,'');           
                $objPHPExcel->setActiveSheetIndex(1)->setCellValue('H' . $i,'');

                $i++;                              
            }
        }    
    }
}

//Inserta linea con datos de cabecera;
$objPHPExcel->setActiveSheetIndexByName('sheet');
$objPHPExcel->getActiveSheet()->insertNewRowBefore(1);
$objPHPExcel->getActiveSheet()->setCellValue('A1','CLASE');
$objPHPExcel->getActiveSheet()->setCellValue('B1','TIPO');
$objPHPExcel->getActiveSheet()->setCellValue('C1','SERIE');
$objPHPExcel->getActiveSheet()->setCellValue('D1','NALBA');
$objPHPExcel->getActiveSheet()->setCellValue('E1','CODIGO');
$objPHPExcel->getActiveSheet()->setCellValue('F1','DESCR');
$objPHPExcel->getActiveSheet()->setCellValue('G1','CANTIDAD');
$objPHPExcel->getActiveSheet()->setCellValue('H1','IMPORTE');


$objPHPExcel->removeSheetByIndex(0);
$objPHPExcel->removeSheetByIndex(1);

//Guardar las modificaciones 
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('xls/Tfrance/LinTuiF.xlsx');


?>
<h2><a href="index_1.php">Volver a Index</a></h2>
</body>
                 
</html>
