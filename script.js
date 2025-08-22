/* <a id="gotop" class="btn btn-warning" href="#" title="Ir arriba">↑</a>

<script>
	/*
	* Script para poder mostrar los títulos de los productos de una forma correcta en las plantillas de TiendaNube.
	* Evitamos el recorte innecesario del título.
	* Favorecemos la experiencia del usuario para buscar lo que necesite, ya sea siendo cliente o vendedor del producto.
	* Mostrar el código del producto en una leyenda aparte.
	* El enlace no se modifica.
	* 
*/

// Función para modificar los títulos de los productos y mostrar los códigos
let acumulador = 0;
let titulares;

function completador(items) {

	// Constante que recoje todos los items de productos.
	titulares = document.querySelectorAll(".js-item-name.item-name");
	console.log("Primera carga: " + titulares.length, "\n", "Acumulador: " + acumulador, "\n", "\n" );


	// Variables a usar para el código del producto en cuestión.
	let codigo;
	let primerParentesis;
	let ultimoParentesis;

	// Agrega un estilo para evitar el recorte en modo móvil
	const estilo = document.createElement("style");
	estilo.textContent = `
		.en-bloques
		{
			display: block;
			margin: -1.5em auto -1.5em auto;
		}
		@media (max-width:769px) {
			.js-item-name.item-name {
				display: block !important;
			}
		}`;
	document.head.appendChild(estilo);
	
	// Recorre cada item del arreglo
	for( i = acumulador; i < titulares.length; i++ )
	{
		// Ubicación del primer paréntesis.
		primerParentesis = titulares[i].title.indexOf("(");
		
		// Ubicación del último paréntesis.
		ultimoParentesis = titulares[i].title.indexOf(")");
		
		// Se recorta el titular sin el código, ni paréntesis, ni espacios.
		titulares[i].innerHTML = titulares[i].title.slice(0, primerParentesis - 1);
		
		// Extracción del código del producto.
		codigo = titulares[i].title.slice( primerParentesis + 1, ultimoParentesis );
		
		// Creamos elementos adicionales
		const br1 = document.createElement("br");
		const small = document.createElement("small");
		const br2 = document.createElement("br");
		
		// Completamos el código con leyenda y formato
		small.textContent = "Código: " + codigo;
		small.classList.add("text-capitalize", "en-bloques");
		
		// Y agregamos los elementos adicionales al DOM
		titulares[i].parentNode.appendChild(br1);
		titulares[i].parentNode.appendChild(small);
		titulares[i].parentNode.appendChild(br2);

	}

	// Actualizamos el valor del acumulador
	acumulador = titulares.length;
	console.log( "Nueva carga de titulares: " + titulares.length, "\n", "Nuevo Acumulador: " + acumulador, "\n", "Diferencia: " + ( titulares.length - acumulador ), "\n", "\n" );

}

// Se ejecuta una vez cargado el DOM.
document.addEventListener("DOMContentLoaded", completador);


// Función para detectar la carga de contenido dinámico
function detectarCargaAjax() {

	// Escuchar el evento scroll para detectar cuando el usuario se desplaza hacia abajo en la página
	window.addEventListener( "scroll", detectarScroll );
	window.addEventListener( "resize", detectarScroll );

	// Altura del footer
	const futer = document.querySelector(".js-hide-footer-while-scrolling.js-footer");

	function detectarScroll() {
		// Verificar si el usuario ha llegado al final de la página
		if (window.innerHeight + window.scrollY >= (document.body.offsetHeight - futer.offsetHeight )) {
			completador();
            ofertor();
		}
	}

	/* IR ARRIBA */
	const toTop = (() => {
		let button = document.getElementById("gotop");
		window.onscroll = () => {
			button.classList[(document.documentElement.scrollTop > 200) ? "add" : "remove"]("is-visible")
		}
		button.onclick = () => {
			console.log( "subiendo..." );
			window.scrollTo( { top:0, behavior:"smooth"} )
		}
	})();

	/* Saltos de línea */
	// Buscar el elemento con la clase especificada
	let alerta = document.querySelector(".shipping-calculator .shipping-calculator-response .alert-warning");

	if (alerta) {
		// Reemplazar cada punto con un punto seguido de un salto de línea (<br />)
		alerta.innerHTML = alerta.innerHTML.replace(/\./g, '.<br /><br />');
	}

	// Saldos y ofertas
	let ancla = document.querySelectorAll(".desktop-nav-item a");
	ancla.forEach( (e) => {
		
		// Verificar si el enlace tiene el texto "Ofertas y Saldos"
		if (e.textContent.trim() === "Ofertas y Saldos") {
		
			// Cambiar el texto del enlace a "Saldos"
			const elEnlace = e;
			elEnlace.addEventListener('click', function(ev) {
				ev.preventDefault();
				elEnlace.href = "#";
			
				const inputSearch = document.getElementsByName('q');
				const searchForm = document.getElementsByClassName('js-search-container js-search-form')[0];
			
				console.log(inputSearch, searchForm);
				// Insertar directamente el texto deseado
				
				inputSearch.forEach((input) => {
					input.value = 'Saldo';
				});
			
				// Enviar el formulario automáticamente
				searchForm.submit();
			});
		}
	});
}

// Ejecutar la función detectarCargaAjax() cuando se carga el DOM
document.addEventListener("DOMContentLoaded", detectarCargaAjax);

/* PROMOCIONES */
let itemsPromos;
let porcentaje;
let acumuladorPromos = 0;
let precioTachado;
let precioViejo;
let precioNegrita;
let arregloPrecios = [];

function convertirADecimal(texto) {
    // Buscar el número antes del símbolo %
    let match = texto.match(/(\d+(?:[.,]\d+)?)\s*%/);
    if (match) {
        // match[1] = parte numérica (puede tener coma o punto)
        let numero = match[1].replace(",", "."); // reemplazar coma por punto
        numero = numero.replace("$", "");
        numero = numero.trim();
        return parseFloat(numero); // convertir a número flotante
    }
    return null; // si no encontró nada
}

function ofertor(promos) {
    itemsPromos = document.querySelectorAll(".js-promotion-label-private.item-label.item-label-offer");

    for( i = acumuladorPromos; i < itemsPromos.length; i++ ) {
        // Obteniendo el porcentaje y convirtiendo a número decimal
        porcentaje = itemsPromos[i].querySelector("span").innerText;
        porcentaje = convertirADecimal(porcentaje);

        // Obteniendo el precio tachado
        precioTachado = itemsPromos[i].parentElement.parentElement.parentElement.parentElement.querySelector(".item-price-container").querySelector(".js-compare-price-display.item-price-compare"); 
        arregloPrecios[i] = precioTachado;
        
        // Mostrando el precio tachado
        precioTachado.style.setProperty("display", "inline", "important");

        // Obteniendo el precio viejo
        precioViejo = precioTachado.parentElement.parentElement.querySelector(".js-price-display.item-price").innerText;
        precioNegrita = precioViejo;
        precioViejo = precioViejo.replace("$", "");
        precioViejo = precioViejo.replace(".", "");
        precioViejo = precioViejo.trim();
        precioViejo = precioViejo.replace(",", ".");
        precioViejo = parseFloat(precioViejo).toFixed(2);

        // Generando el nuevo precio
        precioNuevo = precioViejo - (precioViejo * (porcentaje / 100)).toFixed(2);
        
        // Formateando el nuevo precio y el viejo precio
        precioNuevo = precioNuevo.toLocaleString("es-AR", {
            style: "currency",
            currency: "ARS",
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        });
        precioNuevo = precioNuevo.toString();

        acumuladorPromos = acumuladorPromos + itemsPromos.length
        console.log("Precio Viejo: " + precioViejo + " -> Precio Reducido: " + precioNuevo);

        // Reemplazando el precio en negrita por el nuevo
        arregloPrecios[i].parentElement.parentElement.querySelector(".js-price-display.item-price").innerText = precioNuevo;

        // Reemplazando el precio tachado por el original
        arregloPrecios[i].innerText = precioNegrita;
    }
}
document.addEventListener("DOMContentLoaded", ofertor);


// </script>
