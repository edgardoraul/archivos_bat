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


	// Controla que se haya cargado.
	console.log("Sitio Cargado");

	// Controla que sólo se active cuando agrega cosas al carrito
	$("#add_to_cart button, .ajax_add_to_cart_button").on("click", function()
	{
		const totalCarrito = $(".ajax_cart_total").html();
		console.log(`Agregado al Carrito por ${totalCarrito}`);
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