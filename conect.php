<?php
$serverName = "192.168.0.100\acua"; //serverName\instanceName
$connectionInfo = array( "Database"=>"NET_AFP2019","UID"=>"comandas1","PWD"=>"Acua2018");
$conn = sqlsrv_connect( $serverName, $connectionInfo);
if( $conn ) {
    }else{
     echo "Conexión no se pudo establecer.<br />";
     die( print_r( sqlsrv_errors(), true));
    }
    

