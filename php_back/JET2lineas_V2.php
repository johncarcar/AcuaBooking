<!DOCTYPE html>
<html>
<head>
	<title>Leer Archivo Excel</title>
</head>
<body>
<h1>RESERVAS - Leer lineas de jet2 Archivo Excel</h1>
<p>Archivo Exportado<p>
<?php
//include 'PHPExcel_1.8/Classes/PHPExcel/IOFactory.php';
//require 'PHPExcel_1.8/Classes/PHPExcel.php';
// Creamos un objeto PHPExcel
$objPHPExcel = new PHPExcel();
// Leemos un archivo Excel 2007
$objReader = PHPExcel_IOFactory::createReader('Excel2007');
$objPHPExcel = $objReader->load('xls/Jet2/Book1.xlsx');
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

//----NOTA MUY IMPORTANTE - 
//  EL NUMERO DE ADULTOS DEBE ESTAR EN LA COLUMNA 'G' 
// EL NUMERO DE NULLOS DEBE ESTAR EN LA COLUMNA 'H'

    
   $Nalba=$_GET['numero']; //IMPORTANTE VARIABLE QUE SE DEBE CAMBIAR ANTES DE CREAR
    // $Nalba=1499;
    $num=1;
    $linea=1;
    while ($linea<=$numRows){
            //Pocisión de la Linea a importar.
            
            //Inserta la Primera Línea del Albaran que contiene el nombre del cliente.
            $objPHPExcel->getActiveSheet()->setCellValue('A'.$num,'V');
            $objPHPExcel->getActiveSheet()->setCellValue('B'.$num,'A');
            $objPHPExcel->getActiveSheet()->setCellValue('C'.$num,'AV');
            $objPHPExcel->getActiveSheet()->setCellValue('D'.$num,$Nalba);
            $objPHPExcel->getActiveSheet()->setCellValue('E'.$num,'DESC');
            //inserto el valor de la celda a grabar en una variable.
            $Desc =$objPHPExcel->getActiveSheet()->getCell('M'.$linea);
            $objPHPExcel->getActiveSheet()->setCellValue('F'.$num,$Desc);
            $objPHPExcel->getActiveSheet()->setCellValue('G'.$num,'');
            $objPHPExcel->getActiveSheet()->setCellValue('H'.$num,'');
            $num++;
            
// ---- ACOMPAÑANTES - Se realizarán 2 lineas, ya que muchas viene con mas de 255 caracteres
            //Inserta la Primera Línea del Albaran que contiene los Acompañantes.
            
            $Acomp =$objPHPExcel->getActiveSheet()->getCell('N'.$linea);
                //Qquitamos los caracteres que no nos interezan de los cAcompañantes
                $Acomp= str_replace('One Bedroom apartment',' ',$Acomp);
                $Acomp= str_replace('Two Bedroom apartment',' ',$Acomp);
                $Acomp= str_replace('Deluxe Double room',' ',$Acomp);
                $Acomp= str_replace('Double room',' ',$Acomp);
                $Acomp= str_replace('with Pool View',' ',$Acomp);
                $Acomp= str_replace('for Sole Use',' ',$Acomp);

            $longitud= strlen($Acomp);
            $resto=$longitud-255;
            $AcompDer= substr($Acomp,0,255);
            
            //Inserta la Segunda Línea del Albaran que contiene los Acompañantes.           
            $objPHPExcel->getActiveSheet()->setCellValue('A'.$num,'V');
            $objPHPExcel->getActiveSheet()->setCellValue('B'.$num,'A');
            $objPHPExcel->getActiveSheet()->setCellValue('C'.$num,'AV');
            $objPHPExcel->getActiveSheet()->setCellValue('D'.$num,$Nalba);
            $objPHPExcel->getActiveSheet()->setCellValue('E'.$num,'ACOMP');
            //inserto el valor de la celda a grabar en una variable.
            $objPHPExcel->getActiveSheet()->setCellValue('F'.$num,$AcompDer);
            $objPHPExcel->getActiveSheet()->setCellValue('G'.$num,'');
            $objPHPExcel->getActiveSheet()->setCellValue('H'.$num,'');
            $num++;            
            
            if ($resto>0){
                    // echo '<br> - resto - '.$resto;
                    $AcompIzq= substr($Acomp,-$resto);
                 
                    //Inserta la Segunda Línea del Albaran que contiene los Acompañantes.
                    $objPHPExcel->getActiveSheet()->setCellValue('A'.$num,'V');
                    $objPHPExcel->getActiveSheet()->setCellValue('B'.$num,'A');
                    $objPHPExcel->getActiveSheet()->setCellValue('C'.$num,'AV');
                    $objPHPExcel->getActiveSheet()->setCellValue('D'.$num,$Nalba);
                    $objPHPExcel->getActiveSheet()->setCellValue('E'.$num,'ACOMP');
                    //inserto el valor de la celda a grabar en una variable.
                    $Acomp =$objPHPExcel->getActiveSheet()->getCell('N'.$linea);
                    $objPHPExcel->getActiveSheet()->setCellValue('F'.$num,$AcompIzq);
                    $objPHPExcel->getActiveSheet()->setCellValue('G'.$num,'');
                    $objPHPExcel->getActiveSheet()->setCellValue('H'.$num,'');
                    $num++;
            } //FIN DE CICLO IF            
           
                        
                $persona=0;
                $suma=0;
                $adulto=0;
                $nino=0;
                $sp="<br>";//espacios abreviados

                for ($index=0;$index<(strlen($Acomp));$index++){
                    $persona=(explode('/',$Acomp));
                    $array = preg_split('/[^0-9]+/i', $Acomp);
                    }      
                    for ($index=0;$index<count($array);$index++){
                        $edad= $array[$index];
                        if($edad>11) {
                           //echo $array[$index]; // MUestra cada valor del Arreglo en la pocisión Index.
                           //echo " = Es un Adulto - ";
                           $adulto++;
                        }               
                        if(($edad<12) and ($edad>3)){
                           //echo $array[$index]; // MUestra cada valor del Arreglo en la pocisión Index.
                          // echo " = Es un Niño - ";
                           $nino++;
                    }   
                        
                }
                //echo " ---> Albaran Nº $Nalba // Total Adultos = $adulto / Totla Niños = $nino $sp$sp$sp";
            

                //Inserta la Primera Línea del Albaran que contiene Cantidad ADULTOS
                $objPHPExcel->getActiveSheet()->setCellValue('A'.$num,'V');
                $objPHPExcel->getActiveSheet()->setCellValue('B'.$num,'A');
                $objPHPExcel->getActiveSheet()->setCellValue('C'.$num,'AV');
                $objPHPExcel->getActiveSheet()->setCellValue('D'.$num,$Nalba);
                $objPHPExcel->getActiveSheet()->setCellValue('E'.$num,'7770200014');
                //inserto el valor de la celda a grabar en una variable.
                $objPHPExcel->getActiveSheet()->setCellValue('F'.$num,'JET2 UK ADULTO 18 ORI');
                $objPHPExcel->getActiveSheet()->setCellValue('G'.$num,$adulto);
                $objPHPExcel->getActiveSheet()->setCellValue('H'.$num,'52,00');       
                $num++;

                
                
                
                //Inserta la Primera Línea del Albaran que contiene Cantidad NIÑOS
                $objPHPExcel->getActiveSheet()->setCellValue('A'.$num,'V');
                $objPHPExcel->getActiveSheet()->setCellValue('B'.$num,'A');
                $objPHPExcel->getActiveSheet()->setCellValue('C'.$num,'AV');
                $objPHPExcel->getActiveSheet()->setCellValue('D'.$num,$Nalba);
                $objPHPExcel->getActiveSheet()->setCellValue('E'.$num,'7770200013');
                //inserto el valor de la celda a grabar en una variable.
                $objPHPExcel->getActiveSheet()->setCellValue('G'.$num,$nino);
                $objPHPExcel->getActiveSheet()->setCellValue('F'.$num,'JET2 UK NIÑO 18 ORI');
                $objPHPExcel->getActiveSheet()->setCellValue('G'.$num,$nino);
                $objPHPExcel->getActiveSheet()->setCellValue('H'.$num,'25,10');       
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
$objWriter->save("xls/jet2/LINjet.xlsx");


$Nalba--;
echo'<br> El último albaran creado es el: '.$Nalba.'<br>';
?>


</body>
                 
</html>
