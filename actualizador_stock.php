<?php 
// El silencio es Adamantium o Vibranium (más valioso que el oro).

/* Actualización de la parte del stock */

//Conexion con la base
include "/home/rerda/public_html/config/settings.inc.php";

$con = @mysqli_connect( _DB_SERVER_, _DB_USER_, _DB_PASSWD_, _DB_NAME_ );

// Importando el archivos
$info = fopen ( "/home/rerda/public_html/upload/stock.csv" , "r" );
while ( ( $datos = fgetcsv( $info, 10000, ",") ) !== FALSE )
{
	$linea[] = array( 'Referencia' => $datos[0], 'Cantidad' => $datos[1] );
}
fclose ( $info );

$insertados		= 0;
$errores 		= 0;
$actualizados 	= 0;

// Vaciamos la tabla antes que nada
mysqli_query( $con, "TRUNCATE TABLE importacion_stock;" );

// Ahora actualizaremos los campos
foreach($linea as $indice=>$value)
{

	$codigo = $value["Referencia"];
	$campo2 = $value["Cantidad"];


	//Creamos la sentencia SQL y la ejecutamos
	$sql = mysqli_query( $con, "SELECT * FROM importacion_stock WHERE Referencia = '$codigo'" );
	$num = mysqli_num_rows( $sql );

	if ( $num == 0 )
	{
		$sql = "INSERT INTO importacion_stock (Referencia, Cantidad) VALUES( '$codigo', '$campo2' )";
		
		if ( $insert = mysqli_query( $con, $sql ) )
		{
			$insertados += 1;
		} else {
			$errores += 1;
		}
	}
	
	/*
	// Comento esto porque me interesa que la tabla esté vacía primero que nada. Es más rápido.
	else
	{
		$sql = "UPDATE importacion_stock SET Cantidad = '$campo2' WHERE Referencia = '$codigo'";
		
		if ( $update = mysqli_query( $con, $sql ) )
		{
			$actualizados += 1;
		} else {
			$errores += 1;
		}
	}
	*/
}
/*
	Esta dos consultas van a generar un error con stock de más o de menos en las publicaciones que tienen combinaciones. Pero en el resto se supone que no.

// Primera consulta
mysqli_query( $con, "UPDATE ps_product, importacion_stock SET `ps_product`.`quantity` = 1 WHERE `ps_product`.`reference` = `importacion_stock`.`Referencia`" );

// Segunda consulta
mysqli_query( $con, "UPDATE ps_stock_available, ps_product SET `ps_stock_available`.`quantity` = `ps_product`.`quantity` WHERE `ps_stock_available`.`id_product` = `ps_product`.`id_product`" );

/* 
	Estas consultas corrigen los errores arrastrados por las dos consultas anteriores,
	en cuanto al stock de los productos que tienen combinaciones.
*/

// Tercera consulta
mysqli_query( $con, "UPDATE ps_product_attribute, importacion_stock SET `ps_product_attribute`.`quantity` = `importacion_stock`.`Cantidad` WHERE `ps_product_attribute`.`reference` = `importacion_stock`.`Referencia`" );

// Cuarta constulta
mysqli_query( $con, "UPDATE ps_stock_available, ps_product_attribute SET `ps_stock_available`.`quantity` = `ps_product_attribute`.`quantity` WHERE `ps_stock_available`.`id_product_attribute` = `ps_product_attribute`.`id_product_attribute`" );

// Control de salida
echo "Registros insertados: " . number_format( $insertados, 2 ) . " <br/>";
echo "Registros actualizados: " .number_format( $actualizados, 2 ) . " <br/>";
echo "Errores: " . number_format( $errores, 2 ) . " <br/>";

?>