<!DOCTYPE html>
<html>
<head>
	<title>Leer Archivo Excel</title>
</head>
<body>
<h1>RESERVAS - Leer lineas de IBEROSERVICE - MARISMAS -  Archivo Excel</h1>

<?php
//include 'PHPExcel_1.8/Classes/PHPExcel/IOFactory.php';
//require 'PHPExcel_1.8/Classes/PHPExcel.php';
// Creamos un objeto PHPExcel
$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel5');
$objPHPExcel = $objReader->load('xls/Iberoservice/marismas_1.xls');
// Indicamos que se pare en la hoja uno del libro
$objPHPExcel->setActiveSheetIndex(0);
$allcell = $objPHPExcel->setActiveSheetIndex()->getHighestRow();

$nueva=new PHPExcel_Worksheet($objPHPExcel,'sheet2');
$objPHPExcel->addSheet($nueva,1);
 
 //inicia nuevas variables
// variables a utilizar con sus valores:

//BAJA 
    $codPrecNS='25';
    $codPrecA='55';
    $codArtA='7770200060';
    $codArtNS='7770200059';
    $descArtA='T COOK UK ADULTO 19 BAJA';
    $descArtNS='T COOK UK NI/SEN 19 BAJA';


$Nalba= $_GET['numero'];

$j=1;
$varold=''; 

for ($i=1;$i<=$allcell;$i++){
    $varM= $objPHPExcel->setActiveSheetIndex(0)->getCell('A'.$i); // BOOKING
    $varV= $objPHPExcel->setActiveSheetIndex(0)->getCell('G'.$i); // NOMBRE
    $varAN= $objPHPExcel->setActiveSheetIndex(0)->getCell('U'.$i); // EDAD
    
    //VALOR DE LA EDAD
    if (($varAN<'4' and $varAN<>'')) {//NIÑO
       // echo 'infante, no tomar en cuenta <br>';
        
        }    
    /*if($varAN>='60' and $varAN<>''){// AL PRINCIPIO SE APLICABA OTRO PRECIO
        $varDesc=$descArtA;
        $varArt=$codArtA;
        $varPrec=$codPrecA;  
        }*/
    if($varAN>'11'){//ADULTO
        $varDesc=$descArtA;
        $varArt=$codArtA;
        $varPrec=$codPrecA;        
        }
    if($varAN>'3' and $varAN<'12'){//NIÑO
        $varDesc=$descArtNS;
        $varArt=$codArtNS;
        $varPrec=$codPrecNS;
        }

        if ($varAN>='4'){    
        //DESCARTE DE LOS INFANTES        
            //caso 1
            if ($varM<>'' and $varAN<>'' and $varM <>'E.S.' and $varM<> '35660'){
                if ($varM <> $varold){
                    //linea articulo
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('A'.$j,'V');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('B'.$j,'A');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,'AV');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('E'.$j,$varArt);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('F'.$j,$varDesc);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,'1');           
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('H'.$j,$varPrec);
                    // si quiees ka edad -> $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('I'.$j,$varAN);
                    // echo 'V - A - ALB -'.$Nalba.'-'.$varArt.'-'.$varDesc.'-'.'1-'.$varPrec.'-'.$varAN.'<br>';
                    //linea nombre cliente
                    $j++;
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('A'.$j,'V');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('B'.$j,'A');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,'AV');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('E'.$j,'DESC');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('F'.$j,$varV);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,'');           
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('H'.$j,'');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('I'.$j,'');               
                    // echo 'V - A - ALB -'.$Nalba.'-'.$varArt.'-'.$varDesc.'-'.'1-'.$varPrec.'-'.$varV.'<br>';
                    // muestra las variables por pantalla
                    $j++;
                    $varold=$varM;
                    $Nalba++;



                }else{

                    $Nalba--;
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('A'.$j,'V');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('B'.$j,'A');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,'AV');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('E'.$j,$varArt);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('F'.$j,$varDesc);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,'1');           
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('H'.$j,$varPrec);            
                    // si quiees ka edad -> $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('I'.$j,$varAN);
                    // echo 'V - A - ALB -'.$Nalba.'-'.$varArt.'-'.$varDesc.'-'.'1-'.$varPrec.'-'.$varAN.'<br>';
                    //linea nombre cliente
                    $j++;
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('A'.$j,'V');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('B'.$j,'A');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,'AV');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('E'.$j,'DESC');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('F'.$j,$varV);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,'');           
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('H'.$j,'');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('I'.$j,'');
                    // echo 'V - A - ALB -'.$Nalba.'-'.$varArt.'-'.$varDesc.'-'.'1-'.$varPrec.'-'.$varV.'<br>';               
                    // muestra las variables por pantalla
                    $varold=$varM;
                    $j++;
                    $Nalba++;

                }
            }

            //caso 2
            if ($varM=='' and $varAN<>''){
                    $Nalba--;
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('A'.$j,'V');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('B'.$j,'A');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,'AV');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('E'.$j,$varArt);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('F'.$j,$varDesc);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,'1');           
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('H'.$j,$varPrec);            
                    // si quiees ka edad -> $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('I'.$j,$varAN);
                    // echo 'V - A - ALB -'.$Nalba.'-'.$varArt.'-'.$varDesc.'-'.'1-'.$varPrec.'-'.$varAN.'<br>';
                    //linea nombre cliente
                    $j++;
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('A'.$j,'V');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('B'.$j,'A');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('C'.$j,'AV');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('D'.$j,$Nalba);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('E'.$j,'DESC');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('F'.$j,$varV);
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('G'.$j,'');           
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('H'.$j,'');
                    $objPHPExcel->setActiveSheetIndexByName('sheet2')->setCellValue('I'.$j,'');
                    // echo 'V - A - ALB -'.$Nalba.'-'.$varArt.'-'.$varDesc.'-'.'1-'.$varPrec.'-'.$varV.'<br>';
                    // muestra las variables por pantalla                              
                    $j++;
                    $Nalba++;    


            }
        }//fin ciclo  DESCARTE DE LOS INFANTES 
    } 
    //echo '- DESCARTADO '.$varAN.'-';


//Inserta linea con datos de cabecera;
$objPHPExcel->setActiveSheetIndexByName('sheet2');
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
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save("xls/Iberoservice/LINmarismas.xlsx");

$Nalba--;
echo' El último albaran creado es el: '.$Nalba.'<br>';



?>

<p>Archivo Exportado<p>
</body>
                 
</html>
