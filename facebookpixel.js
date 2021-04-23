// Comprueba que se haya cargado JQuery y la carga total de la web
jQuery(document).ready( function()
{
	/******* Corrige el error del flete gratis ***********+*/
	const varA = $(".ajax_cart_shipping_cost .unvisible");
	varA.html("A determinar");

	/************* El Facebook Pixel **************/
	!function(f,b,e,v,n,t,s)
	{
		if( f.fbq ) return;
		n = f.fbq = function()
		{
			n.callMethod ? n.callMethod.apply(n, arguments) : n.queue.push(arguments)
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

	fbq('init', '1335485450135411');
	fbq('track', 'PageView');

	// Control de la carga de la web
	console.log("La página cargó.");

	// Definiendo la variable
	let totalCarrito = $(".price.cart_block_total.ajax_block_cart_total").html().replace("$ ", "").replace(".", "").replace(",", ".").replace(" ", "");
	totalCarrito = parseFloat(totalCarrito).toFixed(2) * 1;


	// Validando sólo si hay 
	if ( $(".price.cart_block_total.ajax_block_cart_total").html() !== "" )
	{
		// Monto total anterior del carrito de compras. Convertido a número decimal
		console.log( `totalCarrito es ${ typeof(totalCarrito) } = ${totalCarrito} `);
		
		// Comprobando si hubo algún cambio o actualización de los valores
		$(".price.cart_block_total.ajax_block_cart_total").on("change", () =>
		{
			totalCarrito = $(".price.cart_block_total.ajax_block_cart_total").html().replace("$ ", "").replace(".", "").replace(",", ".").replace(" ", "");
			totalCarrito = parseFloat(totalCarrito).toFixed(2) * 1;
			categoria();
		});
	}

	// Muestra en consola
	function mostrarConsola(pre, cant, carr)
	{
		console.log( `Precio es ${typeof(pre)} = ${pre}` );
		console.log( `Cantidad es ${typeof(cant)} = ${cant}` );
		console.log( `Agregado al Carrito por ${carr}` );
	}

	// Mostrar en el tracking del féisbuc
	function husmeador(totalCarr, moneda)
	{
		// Activa el evento del píxel féisbuc
		fbq("track", "AddToCart",
		{
			value: totalCarr,
			currency: moneda
		});
	}

	// Cuando se está en la página de categorías
	if ( $("#category").length > 0 )
	{
		// Cuando se hace click a "Agregar al Carrito"
		$(".button.ajax_add_to_cart_button.btn.btn-default").on("click", () =>
		{
			let precio = $(this).parent().prev().find(".price.product-price").html().replace("$ ", "").replace(".", "").replace(",", ".");
			precio = parseFloat(precio).toFixed(2) * 1;

			// Convierte las cantidades a enteros
			let cantidad = 1;
			
			// Obtiene el total del carrito agrendo lo actual
			let productosAgregando = precio * cantidad;
			totalCarrito = totalCarrito + productosAgregando;

			// Para monitorear que sea correcto.
			mostrarConsola(precio, cantidad, totalCarrito);
			
			// Activa el evento del píxel féisbuc
			husmeador(totalCarrito, "ARS");
		});
	}


	// En la página del producto o en la home
	if ( $("#product").length > 0 || $("#index").length > 0 )
	{
		// Cuando se hace click a "Agregar al Carrito"
		$("#add_to_cart button, .ajax_add_to_cart_button").on("click", () =>
		{
			// Convierte los precios a decimal
			let precio = $("#our_price_display").attr("content").replace("$ ", "");
			precio = parseFloat(precio).toFixed(2);

			// Convierte las cantidades a enteros
			let cantidad = $("#quantity_wanted").val();
			cantidad = parseInt(cantidad);
			
			// Obtiene el total del carrito agrendo lo actual
			let productosAgregando = precio * cantidad;
			totalCarrito = totalCarrito + productosAgregando;

			// Para monitorear que sea correcto.
			mostrarConsola(precio, cantidad, totalCarrito);
			
			// Activa el evento del píxel féisbuc
			husmeador(totalCarrito, "ARS");
		});
	}


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
		const TotalCompra = window.dataLayer.find(monto => monto.transactionTotal);
		TotalCompra2 = TotalCompra.transactionTotal;
		console.log(`La compra ya está hecha por $ ${ TotalCompra2 }`);
		
		fbq("track", "Purchase",
		{
			value: TotalCompra2,
			currency: "ARS"
		});
	}
});



function validador()
{
	const miCbu = document.querySelector("#up_CBU");
	miCbu.addEventListener("blur", validando);
	miCbu.addEventListener("change", validando);
	function validando()
	{
		if( miCbu.value.length === 22)
		{
			console.log("Todo correcto");
		}
		else
		{
			alert("El CBU debe tener 22 dígitos numéricos");
		}
	}
}