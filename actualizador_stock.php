<?php 
/*
	El silencio es Adamantium o Vibranium (más valioso que el oro).
	
	========= Actualización del stock =============
	Versión 1.0
*/

// Incorporamos la información de la conexión
require_once "/home/rerda/public_html/config/settings.inc.php";

// Conexion con la base
$con = @mysqli_connect( _DB_SERVER_, _DB_USER_, _DB_PASSWD_, _DB_NAME_ );

// Importamos el archivo
$info = fopen ( "/home/rerda/public_html/upload/stock.csv" , "r" );
while ( ( $datos = fgetcsv( $info, 10000, "," ) ) !== FALSE )
{
	$linea[] = array( 'Referencia' => $datos[0], 'Cantidad' => $datos[1] );
}
fclose ( $info );

// Variables de control de salida posterior
$insertados 	= 0;
$errores 		= 0;
$actualizados 	= 0;

// Vaciamos las tablas antes que nada. Total sirve para borrar e insertar datos
mysqli_query( $con, "TRUNCATE TABLE importacion_stock;" );
mysqli_query( $con, "TRUNCATE TABLE importacion_stock_acumulado" );


// Apertura de Tabla.
/* echo '<!DOCTYPE HTML>
<!--[if lt IE 7]> <html class="no-js lt-ie9 lt-ie8 lt-ie7" lang="es-es"><![endif]-->
<!--[if IE 7]><html class="no-js lt-ie9 lt-ie8 ie7" lang="es-es"><![endif]-->
<!--[if IE 8]><html class="no-js lt-ie9 ie8" lang="es-es"><![endif]-->
<!--[if gt IE 8]> <html class="no-js ie9" lang="es-es"><![endif]-->
<html lang="es-es">
	<head>
		<meta charset="utf-8" />
		<title>Listado de Stock - Rerda</title>
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
				<th>Stock</th>
			</tr>
		</thead>
	<tbody>"; */

// Ahora actualizaremos los campos
foreach( $linea as $indice => $value )
{
	$codigo = $value["Referencia"];

	// Operación ternaria. Si es menor a 2, coloca 0
	$campo2 = $value["Cantidad"] < 2 ? 0 : $value["Cantidad"];
	// $campo2 = $value["Cantidad"];

	// Creamos la sentencia SQL y la ejecutamos
	$sql = mysqli_query( $con, "SELECT * FROM importacion_stock WHERE Referencia = '$codigo'" );
	$num = mysqli_num_rows( $sql );

	if ( $num == 0 )
	{
		$sql = "INSERT INTO importacion_stock ( Referencia, Cantidad ) VALUES( '$codigo', '$campo2' )";
		
		if ( $insert = mysqli_query( $con, $sql ) )
		{
			$insertados += 1;
		}
		else 
		{
			$errores += 1;
		}


		// Generación de una tabla de stock
		/*echo "<tr><td>".$codigo."</td><td>".$campo2."</td></tr>";*/
	}
}


/* 
	Estas consultas actualizan el stock de los productos que tienen combinaciones.
*/

// El stock de las combinaciones/variantes
mysqli_query( $con, "
	UPDATE ps_product_attribute, importacion_stock 
	SET quantity = `importacion_stock`.`Cantidad` 
	WHERE `ps_product_attribute`.`reference` = `importacion_stock`.`Referencia`;
	"
);

// El stock disponible de las combinaciones/variantes. Se muestra en front y back.
mysqli_query( $con, "
	UPDATE ps_stock_available, ps_product_attribute 
	SET `ps_stock_available`.`quantity` = `ps_product_attribute`.`quantity` 
	WHERE `ps_stock_available`.`id_product_attribute` = `ps_product_attribute`.`id_product_attribute`;
	"
);



/*  ESTAS CONSULTAS VAN A LLENAR Y A RELLENAR CON INFORMACION TABLAS TEMPORALES  */

// Sólo se completa para intercambio de datos. Lo importante es la referencia y el id_product. 
mysqli_query( $con, "
	INSERT INTO importacion_stock_acumulado (id_producto, Referencia, Acumulado)
	SELECT id_product, reference, 0 
	FROM ps_product;
	"
);

// Inserta las cantidades de todos los productos. Los individuales estarán bien. Los que tienen combinaciones/variantes tendrán cero. Pero es temporal.
mysqli_query( $con, "
	UPDATE importacion_stock_acumulado, importacion_stock 
	SET `Acumulado` = `importacion_stock`.`Cantidad` 
	WHERE `importacion_stock_acumulado`.`Referencia` = `importacion_stock`.`Referencia`;
	"
);

// Reemplaza las cantidades totales de los que tienen combinaciones. Va sumando en forma acumulativa.
mysqli_query( $con,
	"
	UPDATE importacion_stock_acumulado, ps_product_attribute
	SET Acumulado = 
		(
		SELECT	SUM(quantity)

		FROM ps_product_attribute
		WHERE id_producto = id_product

		GROUP BY id_product
		)
	WHERE id_producto = id_product;
	"
);

// Actualiza la tabla ps_product ya con los totales definitivos.
mysqli_query( $con, "
	UPDATE ps_product, importacion_stock_acumulado 
	SET `ps_product`.`quantity` = `importacion_stock_acumulado`.`Acumulado` 
	WHERE `ps_product`.`id_product` = `importacion_stock_acumulado`.`id_producto`;
	"
);

// Ultima consulta. Coloca los totales de stock disponibles de todos los productos. Se muestra en el front y el back.
mysqli_query( $con, "
	UPDATE ps_stock_available, importacion_stock_acumulado 
	SET `ps_stock_available`.`quantity` = `importacion_stock_acumulado`.`Acumulado` 
	WHERE `ps_stock_available`.`id_product` = `importacion_stock_acumulado`.`id_producto`;
	"
);

/* LA VOLVEMOS A COLOCAR PARA CORREGIR EVENTUALES ERRORES DE LAS ANTERIORES CONSULTAS */

// El stock de las combinaciones/variantes
mysqli_query( $con, "
	UPDATE ps_product_attribute, importacion_stock 
	SET quantity = `importacion_stock`.`Cantidad` 
	WHERE `ps_product_attribute`.`reference` = `importacion_stock`.`Referencia`;
	"
);

// El stock disponible de las combinaciones/variantes. Se muestra en front y back.
mysqli_query( $con, "
	UPDATE ps_stock_available, ps_product_attribute 
	SET `ps_stock_available`.`quantity` = `ps_product_attribute`.`quantity` 
	WHERE `ps_stock_available`.`id_product_attribute` = `ps_product_attribute`.`id_product_attribute`;
	"
);

/* // Desactiva los productos con stock cero.
mysqli_query( $con, "UPDATE ps_product SET `ps_product`.`active` = 0 WHERE `ps_product`.`quantity` = 0 ");
mysqli_query( $con, "UPDATE ps_product_shop, ps_product SET `ps_product_shop`.`active` = 0 WHERE `ps_product_attribute`.`quantity` = 0 ")



// Activa los productos con stock superior a cero.
mysqli_query( $con, "UPDATE ps_product, ps_product SET `ps_product`.`active` = 1 WHERE `ps_product`.`quantity` > 0 ");
mysqli_query( $con, "UPDATE ps_product_shop, ps_product SET `ps_product_shop`.`active` = 1 WHERE `ps_product_attribute`.`quantity` > 0 ")
 */

// Control de salida
echo "";
echo "<tr><td>Registros insertados: </td><td>" . number_format( $insertados, 2 ) . "</td></tr>";
echo "<tr><td>Registros actualizados: </td><td>"  .number_format( $actualizados, 2 ) . " </td></tr>";
echo "<tr><td>Errores: </td><td>" . number_format( $errores, 2 ) . "</td></tr>";

// Cierre de tabla
/* echo "
		</tbody>
	</table>
	</body>
</html>
"; */

?>