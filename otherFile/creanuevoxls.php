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
$Nalba='500';
$j=1;
$varold='';    
for ($i=1;$i<=$allcell;$i++){
    $varM= $objPHPExcel->setActiveSheetIndex(0)->getCell('M'.$i);
    $varV= $objPHPExcel->setActiveSheetIndex(0)->getCell('V'.$i);
    $varAN= $objPHPExcel->setActiveSheetIndex(0)->getCell('AN'.$i);
    
    //VALOR DE LA EDAD
    if($varAN>='65'){
        $varEDAD='Senior';
        }
    if($varAN>'11' and $varAN<'65'){
        $varEDAD='Adulto';
        }
    if($varAN>'3' and $varAN<'12'){
        $varEDAD='NIÑO';
        }
    
    //caso 1
    if ($varM<>'' and $varAN<>''){
        if ($varM <> $varold){
            echo 'caso 1-1:';
            echo $varM.'-';
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,$varM);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('E'.$j,$varV);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,$varAN);           
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('H'.$j,$varEDAD);
            echo '- j='.$j.'-';
            echo '- i='.$i.'-';
            $j++;
            $varold=$varM;
            echo $Nalba++.'<br>';

        }else{

            $Nalba--;
            echo 'caso 1-2:';
            echo $varM.'-';
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,$varM);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('E'.$j,$varV);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,$varAN);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('H'.$j,$varEDAD);            
            echo '- j='.$j.'-';
            echo '- i='.$i.'-';
            echo $Nalba.'<br>';
            $varold=$varM;
            $j++;
            $Nalba++;
        }
    }
    
    //caso 2
    if ($varM=='' and $varAN<>''){
            $Nalba--;
            echo 'caso 1-2:';
            echo $varold.'-';
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,$varM);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('E'.$j,$varV);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,$varAN);
            $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('H'.$j,$varEDAD);            
            echo '- j='.$j.'-';
            echo '- i='.$i.'-';
            echo $Nalba.'<br>';
            $j++;
            $Nalba++;            
     
    }

}
$objPHPExcel->removeSheetByIndex(0);
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("C:/XAMPP/HTDOCS/BOOKINGS/xls/Iberoservice/LINmarismas.xlsx");
?>

<p>Archivo Exportado<p>
</body>
                 
</html>
