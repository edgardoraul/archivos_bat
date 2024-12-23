<?php 
// El silencio es Adamantium o Vibranium (más valioso que el oro).

/* Actualización de la parte del preshio, loco!! */
include "/home/rerda/public_html/config/settings.inc.php";

//Conexion con la base
$con = @mysqli_connect( _DB_SERVER_, _DB_USER_, _DB_PASSWD_, _DB_NAME_ );

// Importando el archivos
$info = fopen ( "/home/rerda/public_html/upload/price.csv" , "r" );
while ( ( $datos = fgetcsv( $info, 10000, "," ) ) !== FALSE )
{
	$linea[] = array( 'Referencia' => $datos[0], 'Precio' => $datos[1] );
}
fclose ( $info );

$insertados 	= 0;
$errores 		= 0;
$actualizados 	= 0;

// Vaciamos la tabla antes que nada
mysqli_query( $con, "TRUNCATE TABLE importacion_precios;" );

// Ahora actualizaremos los campos
foreach( $linea as $indice => $value )
{
	$codigo = $value["Referencia"];
	$campo2 = $value["Precio"];

	//Creamos la sentencia SQL y la ejecutamos
	$sql = mysqli_query( $con, "SELECT * FROM importacion_precios WHERE Referencia = '$codigo'" );
	$num = mysqli_num_rows( $sql );

	if ( $num == 0 )
	{
		$sql = "INSERT INTO importacion_precios (Referencia, Precio) VALUES( SUBSTRING( '$codigo', 1, 7), '$campo2')";
		
		if ( $insert = mysqli_query( $con, $sql ) )
		{
			$insertados += 1;
		} else {
			$errores += 1;
		}

	}
	/*
		LO COMENTO PORQUE ME QUIERO ASEGURAR DE QUE LA TABLA ESTE PRIMERO VACIA. ES MAS RAPIDO.
	else
	{
		$sql = "UPDATE importacion_precios SET Precio = '$campo2' WHERE Referencia = SUBSTRING( '$codigo', 1,";
		
		if ( $update = mysqli_query( $con, $sql ) )
		{
			$actualizados += 1;
		} else {
			$errores += 1;
		}
	}
	*/
}

// Actualiza el precio de la tabla

mysqli_query( $con, "UPDATE ps_product, importacion_precios SET `ps_product`.`price` = `importacion_precios`.`Precio` WHERE `ps_product`.`reference` = `importacion_precios`.`Referencia`");

mysqli_query( $con, "UPDATE ps_product_shop, ps_product SET `ps_product_shop`.`price` = `ps_product`.`price` WHERE `ps_product_shop`.`id_product` = `ps_product`.`id_product`");


// Desactiva los productos con precio cero.
mysqli_query( $con, "UPDATE ps_product, importacion_precios SET `ps_product`.`active` = 0 WHERE  `ps_product`.`price` = 0 ");
mysqli_query( $con, "UPDATE ps_product_shop, ps_product SET `ps_product_shop`.`active` = 0 WHERE `ps_product_shop`.`price` = 0 ");



// Activa los productos con precio superior a cero.
mysqli_query( $con, "UPDATE ps_product, importacion_precios SET `ps_product`.`active` = 1 WHERE  `ps_product`.`price` > 0 ");
mysqli_query( $con, "UPDATE ps_product_shop, ps_product SET `ps_product_shop`.`active` = 1 WHERE `ps_product_shop`.`price` > 0 ");



// Control de salida
echo "Registros insertados: " . number_format( $insertados, 2 ) . " <br/>";
echo "Registros actualizados: " . number_format( $actualizados, 2 ) . " <br/>";
echo "Errores: " . number_format( $errores, 2 ) . " <br/>";

?>