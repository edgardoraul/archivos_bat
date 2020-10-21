<?php 
// El silencio es Adamantium o Vibranium (más valioso que el oro).

/* Actualización de la parte del stock */

//Conexion con la base
$con=@mysqli_connect("localhost", "rerda_user16", "[N,c@bX~(+Q6)", "rerda_2017");

// Importando el archivos
$info = fopen ("/home/rerda/public_html/upload/stock.csv" , "r" );
while ( ( $datos = fgetcsv( $info, 10000, ", ") ) !== FALSE )
{
	$linea[] = array('Referencia'=>$datos[0],'Cantidad'=>$datos[1]);
}
fclose ($info);

$insertados = 0;
$errores = 0;
$actualizados = 0;
foreach($linea as $indice=>$value)
{

	$codigo=$value["Referencia"];
	$campo2=$value["Cantidad"];


	//Creamos la sentencia SQL y la ejecutamos
	$sql = mysqli_query($con,"select * from importacion_stock where Referencia ='$codigo'");
	$num = mysqli_num_rows($sql);

	if ( $num == 0 )
	{
		$sql="insert into importacion_stock (Referencia, Cantidad) values('$codigo','$campo2')";
		
		if ($insert = mysqli_query($con,$sql))
		{
			$insertados+=1;
		} else {
			$errores+=1;
		}
	} else {
		$sql="update importacion_stock set Cantidad='$campo2' where Referencia='$codigo'";
		
		if ($update = mysqli_query($con,$sql))
		{
			$actualizados+=1;
		} else {
			$errores+=1;
		}
	}
}


// Segunda consulta
$sql = "UPDATE ps_product_attribute, importacion_stock SET  quantity= cantidad WHERE  `ps_product_attribute`.`reference` = `importacion_stock`.`referencia`";

// Tercera constulta
$sql = "UPDATE  ps_stock_available, ps_product_attribute  SET  `ps_stock_available`.`quantity` = `ps_product_attribute`.`quantity` WHERE  `ps_stock_available`.`id_product_attribute` = `ps_product_attribute`.`id_product_attribute`";

// Control de salida
echo "Registros insertados: ".number_format($insertados,2)." <br/>";
echo "Registros actualizados: ".number_format($actualizados,2)." <br/>";
echo "Errores: ".number_format($errores,2)." <br/>";

?>