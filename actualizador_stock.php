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


// Ahora actualizaremos los campos
foreach( $linea as $indice => $value )
{
	$codigo = $value["Referencia"];
	$campo2 = $value["Cantidad"] < 2 ? 0 : $value["Cantidad"];

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


// Control de salida
echo "Registros insertados: " . number_format( $insertados, 2 ) . " <br/>";
echo "Registros actualizados: "  .number_format( $actualizados, 2 ) . " <br/>";
echo "Errores: " . number_format( $errores, 2 ) . " <br/>";

?>