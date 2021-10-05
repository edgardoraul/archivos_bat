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


// Apertura de Tabla.
echo '<!DOCTYPE HTML>
<!--[if lt IE 7]> <html class="no-js lt-ie9 lt-ie8 lt-ie7" lang="es-es"><![endif]-->
<!--[if IE 7]><html class="no-js lt-ie9 lt-ie8 ie7" lang="es-es"><![endif]-->
<!--[if IE 8]><html class="no-js lt-ie9 ie8" lang="es-es"><![endif]-->
<!--[if gt IE 8]> <html class="no-js ie9" lang="es-es"><![endif]-->
<html lang="es-es">
	<head>
		<meta charset="utf-8" />
		<title>Listado de Precios - Rerda</title>
		<meta name="viewport" content="width=device-width, minimum-scale=0.25, maximum-scale=1.6, initial-scale=1.0" />
		<meta name="apple-mobile-web-app-capable" content="yes" />
		<link rel="icon" type="image/vnd.microsoft.icon" href="/img/favicon.ico?1606832549" />
		<link rel="shortcut icon" type="image/x-icon" href="/img/favicon.ico?1606832549" />
	</head>
	<body>
		';
echo "<table>
		<thead>
			<tr>
				<th>Código</th>
				<th>Precio</th>
			</tr>
		</thead>
	<tbody>";

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

		// Generación de una tabla de stock
		/*echo '<tr><td>' . $codigo . '</td><td>' . number_format($campo2, 2, ",", ".") . '</td></tr>';*/
	}
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
echo "<tr><td>Registros insertados: </td><td>" . number_format( $insertados, 2 ) . "</td></tr>";
echo "<tr><td>Registros actualizados: </td><td>"  .number_format( $actualizados, 2 ) . " </td></tr>";
echo "<tr><td>Errores: </td><td>" . number_format( $errores, 2 ) . "</td></tr>";

// Cierre de tabla
echo "
		</tbody>
	</table>
	</body>
</html>
";

?>