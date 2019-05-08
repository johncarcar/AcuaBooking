
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
$objPHPExcel_old = new PHPExcel();
$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load('xls/Jet2/Book1.xlsx');
$objPHPExcel_old = $objReader->load('xls/Jet2/Book1_V17.xlsx');
// Indicamos que se pare en la hoja uno del libro
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel_old->setActiveSheetIndex(0);


$var17=$objPHPExcel_old->setActiveSheetIndex()->getCell('B4');
$var18=$objPHPExcel->setActiveSheetIndex()->getCell('B4');

$value=$var17;

    echo "is_string(";
    var_export($value);
    echo ") = ";
    echo var_dump(is_string($value));


/*
if ($var17==$var18){
    echo 'Variable 17 = '.$var17.'<br>';
    echo 'Variable 18 = '.$var18.'<br>';
    echo 'Está igual<br>';
    
} else {
    echo 'No es igual<BR>';
    }
echo 'se ha salido del ciclo<BR>';   
echo 'Variable 17 = '.$var17.'<br>';
echo 'Variable 18 = '.$var18.'<br>';

echo date('l jS \of F Y h:i:s A');

//Guardamos el documento.
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("xls/jet2/LINjet.xlsx");
 * 
 * 
 */
?>


</body>
                 
</html>

