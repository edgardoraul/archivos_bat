<?php
/**
 * Mínimo de Compra
 *
 * Plugin Name: Mínimo de Compra
 * Plugin URI:  https://webmoderna.com.ar
 * Description: Permite configurar un mínimo de cantidad de un producto y también de otro.
 * Version:     0.0.1
 * Author:      WebModerna | Estudio Contable y Agencia Web
 * Author URI:  https://github.com/edgardoraul/
 * License:     GPLv2 or later
 * License URI: http://www.gnu.org/licenses/old-licenses/gpl-2.0.html
 * Text Domain: woocommerce-max-quantity
 * Domain Path: /languages
 * Requires at least: 4.9
 * Tested up to: 5.8
 * Requires PHP: 5.2.4
 *
 * This program is free software; you can redistribute it and/or modify it under the terms of the GNU
 * General Public License version 2, as published by the Free Software Foundation. You may NOT assume
 * that you can use any other version of the GPL.
 *
 * This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without
 * even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 */

if ( ! defined( 'ABSPATH' ) )
{
	exit; // Exit if accessed directly
}

// Permite sólo un producto (o un número predefinido de productos) por categoría en el carrito
add_filter( 'woocommerce_add_to_cart_validation', 'mk_cantidad_permitida_por_categoria_en_carrito', 10, 2 );
function mk_cantidad_permitida_por_categoria_en_carrito( $passed, $product_id)
{

	$max_num_products = 6;// cambia el número máximo de productos permitidos por categoría
	$running_qty = 0;

	$restricted_product_cats = array();

	//Restringe a categoría o categorías particulares por slug
	$restricted_product_cats[] = 'vinos';
	//$restricted_product_cats[] = 'cat-slug-two';// descomenta para activar la restricción en una segunda categoría

	// Obtiene el slug de la categoría de producto actual en un array
	$product_cats_object = get_the_terms( $product_id, 'product_cat' );
	foreach( $product_cats_object as $obj_prod_cat ) $current_product_cats[] = $obj_prod_cat->slug;

	// Itera a través de cada artículo del carrito
	foreach( WC()->cart->get_cart() as $cart_item_key => $cart_item )
	{
		// Restringe el $max_num_products de cada categoría
		if( has_term( $current_product_cats, 'product_cat', $cart_item['product_id'] ) )
		{

			// Restringe el $max_num_products de categorías con productos restringidos
			//if( array_intersect($restricted_product_cats, $current_product_cats) &amp;&amp; has_term( $restricted_product_cats, 'product_cat', $cart_item['product_id'] )) {

			// count(selected category) quantity
			$running_qty += (int) $cart_item['quantity'];

			// No se permiten más de los productos permitidos en el carrito
			if( $running_qty <= $max_num_products )
			{
				wc_add_notice( 
					sprintf( 'Sólo está permitido un mínimo de %s '.( 
						$max_num_products > 1 ? 'productos de esta categoría' : 'producto de esta categoría') 
					. ' en el carrito.',  $max_num_products ), 'error' );
				$passed = true; // no agrega el nuevo producto al carrito
				// Para el loop
				break;
			}
		}
	}
	return $passed;
}
?>