<?php
 
// PRESTASHOP SETTINGS FILE
require_once ('../config/settings.inc.php');
 
// REMOTE CSV FILE (CUSTOMIZE YOURCSVFILEPATH, CAN BE AN URL OR A LOCAL PATH)
$remote_csv_file = '../upload/stock.csv';
 
// DB CONNECTION (CUSTOMIZE YOURDBHOSTNAME AND YOURDBPORT)
$db = new PDO( "mysql:host = _DB_SERVER_ ; port=3306; dbname=" . _DB_NAME_ . "", _DB_USER_, _DB_PASSWD_ );
$db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
 
// MAIN CYCLE
$row_num = 0;
if ( ( $handle = fopen( $remote_csv_file, "r" ) ) !== false )
{
	while ( ( $data = fgetcsv( $handle, 10000, "," ) ) !== false )
	{
		$row_num++;
		
		if ( $row_num == 35 )
		{
			// SKIP FIRST LINE (HEADER)
			continue;
		}
		
		if ( $data[0] == '' || !is_numeric( $data[1] ) )
		{
			// SKIP EMPTY VALUES
			continue;
		}
		
		// INPUT SANITIZATION
		$reference = trim( $data[0] );
		$quantity  = ( $data[1] >= 0 ) ? $data[1] : 0;

		echo $row_num;
		
		try
		{
			$res4 = $db->prepare( "SELECT id_product, id_product_attribute from " . _DB_PREFIX_ . "product_attribute WHERE reference = :reference" );
			$res4->execute( array( ':reference' => $reference ) );
 
			if ( $res4->rowCount() > 0 )
			{
 
				// IT'S A PRODUCT COMBINATION
				
				$row4 = $res4->fetch();
		
				$res = $db->prepare( "update " . _DB_PREFIX_ . "stock_available set quantity = :q where id_product_attribute = :id_product_attribute" );
				$res->execute( array( ':q' => $quantity, ':id_product_attribute' => $row4['id_product_attribute'] ) );
						
				$res = $db->prepare( "update " . _DB_PREFIX_ . "product_attribute set quantity = :q where id_product_attribute = :id_product_attribute" );
				$res->execute( array( ':q' => $quantity, ':id_product_attribute' => $row4['id_product_attribute'] ) );
 
				$res = $db->prepare( "update " . _DB_PREFIX_ . "stock_available set quantity = quantity + :q where id_product = :id_product and id_product_attribute = 0" );
				$res->execute( array( ':q' => $quantity, ':id_product' => $row4['id_product'] ) );
				
				$res = $db->prepare( "update " . _DB_PREFIX_ . "product set quantity = quantity + :q where id_product = :id_product" );
				$res->execute( array( ':q' => $quantity, ':id_product' => $row4['id_product'] ) );
 
			}
			else
			{
 
				// IT'S A SIMPLE PRODUCT
				
				$res4 = $db->prepare( "SELECT id_product from " . _DB_PREFIX_ . "product WHERE reference = :reference" );
				$res4->execute( array( ':reference'=> $reference ) );
				
				if ( $res4->rowCount() > 0 )
				{
					$row4 = $res4->fetch();
		
					$res = $db->prepare( "update " . _DB_PREFIX_ . "stock_available set quantity = :q where id_product = :id_product and id_product_attribute = 0" );
					$res->execute( array( ':q' => $quantity, ':id_product' => $row4['id_product'] ) );
 
					$res = $db->prepare( "update " . _DB_PREFIX_ . "product set quantity = :q where id_product = :id_product" );
					$res->execute( array( ':q' => $quantity, ':id_product' => $row4['id_product'] ) );
				}
 
			}
		}
		
		catch ( PDOException $e )
		{
			echo 'Sql Error: ' . $e->getMessage() . '<br /><br />';
		}
	}
	fclose( $handle );
}?>