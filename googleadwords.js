/******* Google Adwords para todas las páginas en general ********/
window.dataLayer = window.dataLayer || [];
function gtag()
{
	dataLayer.push(arguments);
}
gtag('js', new Date());
gtag('config', 'AW-701646845');

// Condicional sólo para el evento de purchace, cuando termninó la compra. 
window.addEventListener("DOMContentLoaded", googleAds);
function googleAds()
{
	const monitoreo = document.querySelector("#order-confirmation");
	if ( monitoreo != null )
	{
		gtag('event', 'conversion',
		{
			'send_to': 'AW-701646845/pVfoCImM1O8BEP2Pyc4C',
			'transaction_id': ''
		});
	}
}

var precioReducido = ["412002434","412002436","412002438","412002440","412002442","412002444","412002546","412002548","412002550","222060100","222060101","222060102","222060103","222060104","222060105","222060106","222060107","5919165"];

var PrecioNoActualiza = ["820511734","820511735","820511736","820511737","820511738","820511739","820511740","820511741","820511742","820511743","820511744","820511745","820511746","820511747","820511748","820511749"];