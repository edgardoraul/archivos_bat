!function(f,b,e,v,n,t,s)
{
	if( f.fbq ) return;
	n = f.fbq = function()
	{
		n.callMethod ? n.callMethod.apply(n,arguments) : n.queue.push(arguments)
	};
	if(!f._fbq)f._fbq = n;
	n.push = n;
	n.loaded = !0;
	n.version = '2.0';
	n.queue = [];
	t = b.createElement(e);
	t.async = !0;
	t.src = v;
	s = b.getElementsByTagName(e)[0];
	s.parentNode.insertBefore(t,s)
} (window, document, 'script', 'https://connect.facebook.net/en_US/fbevents.js');

/*
	Este no funciona, no se porqué.
	fbq('init', '1335485450135411');
*/
fbq('init', '3681335751927871');
fbq('track', 'PageView');


// Espera que se cargue el DOM para ejecutar la función
window.addEventListener("DOMContentLoaded", fesbukeador);
function fesbukeador()
{

	// Para todas las páginas exepto la orden de compra y su terminación.
	if ( $("#order").length > 0 )

	// Muestra que se haya cargado la web.
	console.log("Sitio Cargado");

	// Definciendo la variable
	let totalCarrito = $(".ajax_cart_total").html().replace("$ ", "").replace(".", "").replace(",", ".").replace(" ", "");

	// Validando sólo si hay 
	if ( $(".ajax_cart_total").html() !== "" )
	{
		// Monto total anterior del carrito de compras. Convertido a número decimal
		totalCarrito = parseFloat(totalCarrito).toFixed(2);
		if( typeof(totalCarrito) === "string" )
		{
			totalCarrito = 0;
		}
		console.log( "totalCarrito es " + typeof(totalCarrito) + " = " + totalCarrito );
	}


	// Cuando se hace click a "Agregar al Carrito"
	$("#add_to_cart button, .ajax_add_to_cart_button").on("click", function()
	{
		// Convierte los precios a decimal
		let precio = $("#our_price_display").attr("content").replace("$ ", "");
		precio = parseFloat(precio).toFixed(2);
		console.log( "Precio es " + typeof(precio) + " = " + precio);

		// Convierte las cantidades a enteros
		let cantidad = $("#quantity_wanted").attr("value");
		cantidad = parseInt(cantidad);
		console.log( "cantidad es " + typeof(cantidad) + " = " + cantidad);
		
		// Obtiene el total del carrito agrendo lo actual

		let productosAgregando = precio * cantidad;
		totalCarrito = totalCarrito + productosAgregando;
		console.log("productosAgregando es " + typeof(productosAgregando));

		// Para monitorear que sea correcto.
		console.log(`Agregado al Carrito por ${ totalCarrito }`);
		
		// Activa el evento del píxel féisbuc
		fbq("track", "AddToCart",
		{
			value: totalCarrito,
			currency: "ARS"
		});
	});

	// Cuando alguien inicia el proceso del pago, o ir hacia la caja.
	if ( $("#order").length > 0 )
	{
		fbq('track', 'InitiateCheckout');
		console.log("Comenzando checkout.");
	}

	// Controla que sólo se cargue cuando finalize la compra.
	if ( $("#order-confirmation").length > 0 )
	{
	    // Extrae la información de un script al final de la página.
	    const TotalCompra = window.dataLayer[0].transactionTotal;
		console.log(`La compra ya está hecha por $ ${ TotalCompra }`);
		
		fbq("track", "Purchase",
		{
			value: TotalCompra,
			currency: "ARS"
		});
	}
}