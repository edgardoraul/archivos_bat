/* Google Adwords para todas las páginas en general */
window.dataLayer = window.dataLayer || [];
function gtag(){
	dataLayer.push(arguments);
}
gtag('js', new Date());
gtag('config', 'AW-701646845');

/* Condicional sólo para el evento de purchace, cuando termninó la compra. */
window.addEventListener("DOMContentLoaded", googleAds);
function googleAds()
{
	if ( $("#order-confirmation").length > 0 )
	{
		gtag('event', 'conversion',
		{
			'send_to': 'AW-701646845/pVfoCImM1O8BEP2Pyc4C',
			'transaction_id': ''
		});
	}
}