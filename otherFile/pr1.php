<?php


    include("PHPExcel_1.8/Classes/PHPExcel.php");
     
    //Creas el objeto
    $objPHPExcel   = new PHPExcel(); //Nuevo objeto excel para crear un archivo
     
    //Aquí puedes modificar algunas propiedades del archivo que será creado
    $objPHPExcel->getProperties()->setCreator("Creador");
    $objPHPExcel->getProperties()->setLastModifiedBy("Ultima modificacion");
    $objPHPExcel->getProperties()->setTitle("Office 2007 XLSX Test Document");
    $objPHPExcel->getProperties()->setSubject("Office 2007 XLSX Test Document");
    $objPHPExcel->getProperties()->setDescription("Test document for Office 2007 XLSX, generated using PHPExcel classes.");
     
    //Con ésta función puedes setear las columnas que irán de título
     
    //Ajustas la celda al tamaño del texto
    foreach( range('A','C') as $letra ){ //Recorremos las letras que iran en nuestro titulo
       $objPHPExcel->getActiveSheet()->getColumnDimension($letra)->setAutoSize(true);
    }
     
    //Seteas los titulos
    $objPHPExcel->setActiveSheetIndex(0);
    $objPHPExcel->getActiveSheet()->SetCellValue('A1', 'Campo1');
    $objPHPExcel->getActiveSheet()->SetCellValue('B1', 'Campo2');
    $objPHPExcel->getActiveSheet()->SetCellValue('C1', 'Campo3');   
     
    //Aqui comenzamos a escribir en el archivo excel, toma en cuenta que si decides poner columnas de titulo, debes empezar apartir del renglon #2, esto puede ir en una iteración, dependiendo de cuantos datos necesites, eso te lo dejo a ti ;)
     
    $c = 2; //Numero de renglón
    $objPHPExcel->setActiveSheetIndex(0);
    $objPHPExcel->getActiveSheet()->SetCellValue('A'.$c, 'Valor de mi campo1');
    $objPHPExcel->getActiveSheet()->SetCellValue('B'.$c, 'Valor de mi campo2');
    $objPHPExcel->getActiveSheet()->SetCellValue('C'.$c, 'Valor de mi campo3');
     
    //El nombre de la hoja en tu archivo excel
    $objPHPExcel->getActiveSheet()->setTitle('Example');
     
    //Creamos el archivo
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    $objWriter->save('/var/www/file/archivo.xls');
    echo "Archivo creado: ";//.$namexls;