function(obj, parse)
{
	var partialProduct = parse(obj);
	
	const precioReducido = [
		"412002434",
		"412002436",
		"412002438",
		"412002440",
		"412002442",
		"412002444",
		"412002546",
		"412002548",
		"412002550",
		"222060100",
		"222060101",
		"222060102",
		"222060103",
		"222060104",
		"222060105",
		"222060106",
		"222060107",
		"5919165"
	];

	/* Deshabilitando esto
	var PrecioNoActualiza = [
		"820511734",
		"820511735",
		"820511736",
		"820511737",
		"820511738",
		"820511739",
		"820511740",
		"820511741",
		"820511742",
		"820511743",
		"820511744",
		"820511745",
		"820511746",
		"820511747",
		"820511748",
		"820511749"
	];
	*/
	
	var opciones = [
		{
			suffix: "-CL",
			multiplierprice: 1.05,
			sumador: 0,
			pricer: 4848
		},
		{
			suffix: "-CL-EG",
			multiplierprice: 1.05,
			sumador: 0, //250 (del camnbio que pidieron) + 450 (del envío gratis)
			pricer: 5352
		},
		{
			suffix: "-PR",
			multiplierprice: 1.1,
			sumador: 0,
			pricer: 5926
		},
		{
			suffix: "-PR-EG",
			multiplierprice: 1.1,
			sumador: 0, //250 (del camnbio que pidieron) + 450 (del envío gratis)
			pricer: 6541
		}
	];


	//    _.toString(partialProduct.sku)
	/* Comento para deshabilitar
	if ( _.includes( PrecioNoActualiza, partialProduct.sku ) )
	{
		return null;
	}
	*/

	var basePrice = obj[ 1 ];

	if ( _.includes( precioReducido, partialProduct.sku ) )
	{
		basePrice *= 0.4
	};

	// Comento esta función para deshabilitar esta función inmovilizadora
	/*
	// Variables y función para inmovilizar el precio de los borcegos
	var sonBorcegos = _.some( ['8205115', '8205116', '8205041'], function(sku)
	{
		return _.includes( partialProduct.sku, sku );
	});

	
	if (sonBorcegos)
	{
		
		var products = opciones.map( function(kit)
		{
			var newObj = _.cloneDeep(p artialProduct );
			newObj.sku = partialProduct.sku + kit.suffix;
			var price = kit.pricer;
			newObj.prices = [
				{
					priceList: "Piezas",
					amount: _.round( price, 0 )
				}
			];
			return newObj;
		});
	}
	else 
	{
	*/		
		var products = opciones.map( function(kit)
		{
			var newObj = _.cloneDeep( partialProduct );
			newObj.sku = partialProduct.sku + kit.suffix;
			var price = ( basePrice * kit.multiplierprice + kit.sumador ) * 1.05;

			//if(!newObj.sku.includes("-EG") && price < 2000){price += 450};
			
			newObj.prices = [
				{
					priceList: "Piezas",
					amount: _.round( price, 0 )
				}
			];

			if ( newObj.prices[0].amount < 2000 && newObj.sku.includes("-EG") )
			{
				newObj.prices[0].amount = newObj.prices[0].amount + 400
			};

			if ( newObj.prices[0].amount < 2000 && !newObj.sku.includes("-EG") )
			{
				newObj.prices[0].amount = newObj.prices[0].amount + 250
			};
			
			if ( newObj.prices[0].amount > 2500 && newObj.sku.includes("-EG") )
			{
				newObj.prices[0].amount = newObj.prices[0].amount + 650
			};
			return newObj;
		});
  
	/*
	// Comentado para deshabilitar la función inmovilizador del precio de los borcegos
// Para controlar aquí cuando cambia
const precio_limite_meli = 3500

Precio Final = precio < precio_limite_meli => (precio * 1.05) + 40.
Precio Final = precio >= precio_limite_meli => (precio * 1.05) 450 + 40

// Clásica: -CL
// Premium: -PR
Precio Final = precio < precio_limite_meli => (precio * 1.23) + 40.
Precio Final = precio >= precio_limite_meli => (precio * 1.23) 450 + 40

	}
	*/
	return products;
}

// Tipo de plubicaciones
const Clasica = "-CL";
const Premium = "-PR";

// Para controlar aquí cuando cambia
const valor_limite_meli = 3500;

// Variable del flete promedio a nuestro cargo
const flete_nuestro = 450;

// Variable de costo fijo de MELI
const costo_fijo_meli = 40;

// Precio obtenido del Lince
let precio_lince;

// Multiplicador publicaciones clásicas
const coeficiente_clasica = 1.05;

// Multiplicador publicaciones premium
const coeficiente_premium = 1.23;

if( sku.sufijo === Clasica )
{
	let precio_final = if( precio_lince < valor_limite_meli ) {
		
		( precio_lince * coeficiente_clasica ) + costo_fijo_meli
		
		} elseif ( precio_lince >= valor_limite_meli ) {
		
			( precio_lince * coeficiente_clasica ) +  costo_fijo_meli + flete_nuestro;
		}
}

if( sku.sufijo === Premium )
{
	let precio_final = if( precio_lince < valor_limite_meli ) {
		
		( precio_lince * coeficiente_premium ) + costo_fijo_meli
		
		} elseif ( precio_lince >= valor_limite_meli ) {
		
			( precio_lince * coeficiente_premium ) +  costo_fijo_meli + flete_nuestro;
		}
}

/*
Sin importar si terminan en -EG
De hecho. Nos gustaría que en forma masiva; todas las que terminen con sku "-EG" se les coloque "flete a cargo del comprador". O lo que es lo mismo "Desactivar flete gratis".

*/

// Ofertas especial
const reduccion_precio = 0.4;
const precioReducido = [
	"412002434",
	"412002436",
	"412002438",
	"412002440",
	"412002442",
	"412002444",
	"412002546",
	"412002548",
	"412002550",
	"222060100",
	"222060101",
	"222060102",
	"222060103",
	"222060104",
	"222060105",
	"222060106",
	"222060107",
	"5919165"
];

if ( sku === precioReducido[i] )
{
	precio_final = precio_final - precio_final * reduccion_precio;
}