<!DOCTYPE html>
<!--
To change this license header, choose License Headers in Project Properties.
To change this template file, choose Tools | Templates
and open the template in the editor.
-->
<html>
    <head>
        <meta charset="UTF-8">
        <title></title>
        <style>
            div { text-align: left;
                  padding:10px ;
                  max-width: 50%;
                  margin-left: 5%;
                  margin-top: 0%;
                  background-color: buttonhighlight;
                  font-size: 15;
                  }
            h1,h2,h3  {color:blue;
                margin-left: 20px;
                margin-top: 20px;
                
                  }
            a, input {margin-left:0px;
               margin-top: 0px;
                }      
        </style>
    </head>
    <body>
        <!-- CONSULTA EL ÚLTIMO NÚMERO DE ALVARAN DE LA BASE DE DATOS-->
        <?php
            require 'conect.php';
            $sql= "SELECT max(Numero) as num FROM [Cabecera Facturacion] where Serie='AV'";
            $result=sqlsrv_query($conn,$sql);
            if ($result===false) { die(print_r(sqlsrv_errors()));}
            $row= sqlsrv_fetch_array($result,SQLSRV_FETCH_ASSOC);
            //$row =sqlsrv_fetch_array($result,SQLSRV_FETCH_ASSOC);
            $var1=$row['num']+1;
        ?>
        
        <h2>Elija el listado que desea, escogiendo el número de albarán correspondiente.</h2>
        <h3><?php echo " El último alvarán generado es: $var1" ?></h3>
        <div>
            <form action="JET2Cabecera.php" method="GET" target="_self">
                <h2>JET2</h2>
                <h3>Introduzca el Nº de Albaran para generar las Cabeceras y Albaranes</h3>
                <input type="text" name="numero" value="<?php echo $var1?>">
                <input type="submit" value="Enviar"><br/><p/>
             </form>
        </div>
        <div>
            <form action="PLACabecera_1.php" method="GET" target="_self" >
                <h2>PLAYAPARK</h2>
                <h3>Introduzca el Nº de Albaran para generar las Cabeceras y Albaranes</h3>
                <input type="text" name="numero" value="<?php echo $var1?>">
                <input type="submit" value="Enviar"><br/><p/>
             </form>
        </div>
        <div>
            <form action="OASCabecera_1.php" method="GET" target="_self">
                <h2>OASIS</h2>
                <h3>Introduzca el Nº de Albaran para generar las Cabeceras y Albaranes</h3>
                <input type="text" name="numero" value="<?php echo $var1?>">
                <input type="submit" value="Enviar"><br/><p/>
             </form> 
        </div>
        <div>
            <form action="MARCabecera_1.php" method="GET" target="_self">
                <h2>MARISMAS</h2>
                <h3>Introduzca el Nº de Albaran para generar las Cabeceras y Albaranes</h3>
                <input type="text" name="numero" value="<?php echo $var1?>">
                <input type="submit" value="Enviar"><br/><p/>
             </form>
        </div>
        <div>
            <form action="TuifCabecera.php" method="GET" target="_self">
                <h2>TUI FRANCE</h2>
                <h3>Introduzca el Nº de Albaran para generar las Cabeceras y Albaranes</h3>
                <input type="text" name="numero" value="<?php echo $var1?>">
                <input type="submit" value="Enviar"><br/><p/>
             </form>
        </div>
    </body>
</html>
