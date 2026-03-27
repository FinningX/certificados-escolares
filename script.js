                                                /*Sistema desarrollado para uso exclusivo de la institución Nuestra Señora de Fátima.
                                                            Queda prohibida su copia o distribución sin autorización.*/

let contadorExcel = 0;
let contadorManual = 0;

//funcion para descargar el PDF generado
function descargarPDF(){
    window.print();
}

/*funcion para leer el Excel*/ 
function procesarExcel(){

    contadorExcel = 0;
    contadorManual = 0;

    const fileInput = document.getElementById("excelFile");
    const file = fileInput.files[0];

    if(!file){
        mostrarMensaje("⚠ Debe seleccionar un archivo Excel antes de generar los certificados.", "error");
        return;
    }

    const reader = new FileReader();

    reader.onload = function(e){

        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data,{type:'array'});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        // ✅ IMPORTANTE: mantiene las columnas aunque estén vacías
        let alumnos = XLSX.utils.sheet_to_json(sheet, {
            defval: ""
        });

        // eliminar filas vacías
        alumnos = alumnos.filter(a => a.nombre && a.nombre.toString().trim() !== "");

        if(alumnos.length === 0){
            mostrarMensaje("⚠ El archivo Excel no contiene registros", "error");
            return;
        }

        const columnasRequeridas = [
            "apellido",
            "nombre",
            "dni",
            "nacimiento",
            "edad",
            "grado",
            "solicitante",
            "nombre_solicitante",
            "dni_solicitante",
            "ante",
            "localidad",
            "genero"
        ];

        // nombres más amigables en el mensaje de error
        const nombresBonitos = {
            apellido: "Apellido",
            nombre: "Nombre",
            dni: "DNI",
            nacimiento: "Fecha de nacimiento",
            edad: "Edad",
            grado: "Grado",
            solicitante: "Solicitante",
            nombre_solicitante: "Nombre del solicitante",
            dni_solicitante: "DNI del solicitante",
            ante: "Repartición",
            localidad: "Localidad",
            genero: "Género"
        };

        const columnasExcel = Object.keys(alumnos[0] || {});

        // VALIDAR COLUMNAS FALTANTES
        const columnasFaltantes = columnasRequeridas.filter(col => !columnasExcel.includes(col));

        if(columnasFaltantes.length > 0){

            const mensajeColumnas = columnasFaltantes
                .map(col => nombresBonitos[col])
                .join("\n");

            mostrarMensaje("⚠ El archivo Excel tiene columnas faltantes:\n\n" + mensajeColumnas, "error");
            return;
        }

        // VALIDAR DATOS FALTANTES
        const erroresAgrupados = {};

        alumnos.forEach(alumno => {

            const nombreCompleto = `${alumno.apellido || ""} ${alumno.nombre || ""}`.trim() || "(Sin nombre)";

            columnasRequeridas.forEach(col => {

                const valor = alumno[col];

                if(valor === undefined || valor === null || valor.toString().trim() === ""){

                    if(!erroresAgrupados[nombreCompleto]){
                        erroresAgrupados[nombreCompleto] = [];
                    }

                    erroresAgrupados[nombreCompleto].push(nombresBonitos[col]);
                }

            });

        });

        if(Object.keys(erroresAgrupados).length > 0){

            let mensaje = "⚠ El archivo tiene datos faltantes:\n\n";

            for(let alumno in erroresAgrupados){
                mensaje += `• ${alumno}: ${erroresAgrupados[alumno].join(", ")}\n`;
            }

            mostrarMensaje(mensaje);
            return;
        }

        //SI TODO ESTA OK → generar los certificados
        generarCertificados(alumnos);

    };

    reader.readAsArrayBuffer(file);
}

/*Funcion que genera el pdf*/
function generarPDF(nombreAlumno){

    limpiarCertificadosVacios(); // <-- limpia páginas vacías

        const elemento = document.querySelector(".plantilla");

        const opciones = {

            margin:0,

            filename:`Certificado_${nombreAlumno}.pdf`,

            image:{type:'jpeg',quality:1},

            html2canvas:{
                scale:3,
                useCORS:true,
                letterRendering:true
            },

            jsPDF:{unit:'mm',format:'a4',orientation:'portrait'}

    };

    return html2pdf().set(opciones).from(elemento).save();

}

//funcion para generar un PDF con varios certificados (todos los que estén en el contenedor)
function generarPDFMultiple(){

    limpiarCertificadosVacios(); // <-- limpia páginas vacías

    const elemento = document.getElementById("contenedorPDF");

    const opciones = {

        margin:0,
        filename:"certificados_alumnos.pdf",
        image:{type:'jpeg',quality:1},
        html2canvas:{
                scale:3,
                useCORS:true,
                letterRendering:true
            },
        jsPDF:{unit:'mm',format:'a4',orientation:'portrait'}

    };

    html2pdf().set(opciones).from(elemento).save();

}

//funcion para convertir la fecha de Excel al formato 00/00/0000, teniendo en cuenta que Excel maneja las fechas como números de serie
function convertirFechaExcel(fecha){

    if(typeof fecha === "number"){

        const utc_days  = Math.floor(fecha - 25569);
        const utc_value = utc_days * 86400;                                        
        const fechaJS = new Date(utc_value * 1000);

        const dia = String(fechaJS.getUTCDate()).padStart(2,'0');
        const mes = String(fechaJS.getUTCMonth() + 1).padStart(2,'0');
        const anio = fechaJS.getUTCFullYear();

        return `${dia}/${mes}/${anio}`;
    }

    return fecha;
}

/*funcion que genera los certificados automaticamente*/ 
async function generarCertificados(alumnos){

    const contenedor = document.getElementById("contenedorPDF");

    contenedor.innerHTML = "";

    const modelo = document.querySelector(".plantilla");

    let i = 0;
    try {
    for(let alumno of alumnos){

        i++;
        mostrarMensaje(`Generando certificados... ${i}/${alumnos.length}`);

        const copia = modelo.cloneNode(true);
        copia.classList.remove("plantilla");
        aplicarGenero(copia, alumno.genero);
        /*----------------------*/
        const nombreCompletoAlumno = 
        `${alumno.apellido || ""} ${alumno.nombre || ""}`.trim();

        copia.querySelector(".campoAlumno").innerText = nombreCompletoAlumno;
        /*----------------------*/
        
        /*aplica separador de miles al dni y elimina comas y espacios*/
        alumno.dni = alumno.dni.toString().replace(/\D/g, "");
        copia.querySelector("#dni").value = Number(alumno.dni).toLocaleString("es-AR");

        /*Convierte la fecha a formato 00/00/0000*/
        copia.querySelector("#nacimiento").value = convertirFechaExcel(alumno.nacimiento);

        copia.querySelector("#edad").value = alumno.edad;
        
        //formatea el grado y agrega el simbolo de grado
        copia.querySelector("#grado").value =  formatearGrado(alumno.grado);

        copia.querySelector("#solicitante").value = alumno.solicitante;

        //formatea el nombre del solicitante eliminando caracteres no permitidos y espacios al inicio y final
        copia.querySelector("#nombre_solicitante").value =
        formatearNombreSolicitante(alumno.nombre_solicitante)
            .toLocaleUpperCase("es-AR");

        /*aplica separador de miles al dni y elimina comas y espacios*/
        alumno.dni_solicitante = alumno.dni_solicitante.toString().replace(/\D/g, "");
        copia.querySelector("#dni_solicitante").value = Number(alumno.dni_solicitante).toLocaleString("es-AR");

        /*--------------------------------------------------------------------------------------------------------*/

        let textoAnte = (alumno.ante || "").trim();

        let textoLinea1 = textoAnte;
        let textoLinea2 = "";

        const limite = 50;

        if(textoAnte.length > limite){

            let corte = textoAnte.lastIndexOf(" ", limite);

            if(corte === -1) corte = limite;

            textoLinea1 = textoAnte.substring(0, corte).trim();
            textoLinea2 = textoAnte.substring(corte).trim();
        }

        copia.querySelector(".campoSolicitante").innerText = textoLinea1;
        copia.querySelector(".campoSolicitante2").innerText = textoLinea2;

        const linea2p = copia.querySelector(".campoSolicitante2").closest("p");

        if(!textoLinea2){
            linea2p.style.display = "none";
        }else{
            linea2p.style.display = "flex";
        }

        /*--------------------------------------------------------------------------------------------------------*/

        copia.querySelector("#localidad").value = alumno.localidad?.trim();
        if(i < alumnos.length){
            copia.style.pageBreakAfter = "always";
        }else{
            copia.style.pageBreakAfter = "auto";
        }

        contenedor.appendChild(copia);
        await new Promise(r => setTimeout(r,10)); 
    } } catch(error){
    console.error("Error generando certificado:", error);
    mostrarMensaje("Error generando certificado. Revisar consola.");
}
    contadorExcel += alumnos.length;

    mostrarMensaje(
        `✔ ${contadorExcel} certificados generados desde Excel
        ✏ ${contadorManual} certificados agregados manualmente`
    );
    /*generarPDFMultiple();*/
}

//funcion para formatear el grado y agregar el simbolo de grado
function formatearGrado(valor){

    if(valor === undefined || valor === null) return "";

    let texto = String(valor).trim();

    if(texto === "") return "";

    // eliminar todo lo que no sea número
    let numero = texto.replace(/\D/g, "");

    if(numero === "") return "";

    return numero + "°";
}

//funcion para dividir el texto largo en dos lineas dentro del certificado
function dividirTextoEnLineas(texto, campo1, campo2){

    const palabras = texto.split(" ");

    let linea1 = "";
    let linea2 = "";

    campo1.innerText = "";
    campo2.innerText = "";

    for(let palabra of palabras){

        let prueba = (linea1 + " " + palabra).trim();

        campo1.innerText = prueba;

        if(campo1.scrollWidth > campo1.clientWidth){
            linea2 += palabra + " ";
        }else{
            linea1 = prueba;
        }

    }

    campo1.innerText = linea1.trim();
    campo2.innerText = linea2.trim();

}



/*Funcion para detectar el genero y tachar Dn o Dña*/
function aplicarGenero(certificado, genero){

    const dn = certificado.querySelector("#dn");
    const dna = certificado.querySelector("#dna");

    dn.style.textDecoration = "none";
    dna.style.textDecoration = "none";

    if(!genero) return;

    genero = genero.toString().trim().toUpperCase();

    // si el género es femenino, tachar "Dn" y dejar "Dña" detectando tanto "F" como "Femenino" ya sea en mayúscula o minúscula. Si el género es masculino, tachar "Dña" y dejar "Dn" detectando tanto "M" como "Masculino" 
    if(genero === "F" || genero === "FEMENINO"){
        dn.style.textDecoration = "line-through";
    }

    if(genero === "M" || genero === "MASCULINO"){
        dna.style.textDecoration = "line-through";
    } 

}

//funcion para eliminar certificados vacios antes de exportar y descargar el PDF
function limpiarCertificadosVacios(){

    const contenedor = document.getElementById("contenedorPDF");

    const certificados = contenedor.querySelectorAll(".certificado");

    certificados.forEach(cert => {

        const nombre = cert.querySelector(".campoAlumno")?.innerText.trim();

        if(!nombre){
            cert.remove();
        }

    });

}

/*funcion para agregar certificado manualmente si falto alguno*/
function agregarCertificadoManual(){

const contenedor = document.getElementById("contenedorPDF");

const plantilla = document.querySelector(".plantilla");

const copia = plantilla.cloneNode(true);

const generoManual = document.getElementById("genero_manual").value;

aplicarGenero(copia, generoManual);

copia.classList.remove("plantilla");

/* detectar si el segundo renglon esta vacio */

const ante2 = copia.querySelector(".campoSolicitante2").innerText.replace(/\s+/g,'').trim();

const linea2 = copia.querySelector(".campoSolicitante2").closest("p");

if(!ante2){
    linea2.style.display = "none";
}else{
    linea2.style.display = "flex";
}

/* bloquear edicion */

/*copia.querySelectorAll("[contenteditable]").forEach(el=>{
    el.contentEditable = "false";
});

copia.querySelectorAll("input").forEach(el=>{
    el.readOnly = true;
});*/

contenedor.appendChild(copia);

/* limpiar plantilla */

plantilla.querySelectorAll("input").forEach(el=> el.value="");

plantilla.querySelectorAll("[contenteditable]").forEach(el=> el.innerText="");

contadorManual++;

mostrarMensaje(
`✔ ${contadorExcel} certificados generados desde Excel
✏ ${contadorManual} certificados agregados manualmente`
);

document.getElementById("genero_manual").value = "";
}

/*funcion para mostrar mensajes de estado*/
function mostrarMensaje(texto, tipo="ok"){

    const mensaje = document.getElementById("mensajeEstado");

    if(!mensaje) return;

    mensaje.innerHTML = texto.replace(/\n/g,"<br>");

    mensaje.className = "";

    mensaje.classList.add(tipo === "ok" ? "mensaje-ok" : "mensaje-error");

    mensaje.style.display = "block";

    setTimeout(()=>{
        mensaje.style.display="none";
    },5000);

}


//funcion para eliminar un certificado manualmente agregado
function eliminarCertificado(boton){

    const certificado = boton.closest(".certificado");

    if(confirm("¿Eliminar este certificado?")){
        certificado.remove();
    }

}

//funcion para formatear el dni con separador de miles mientras se escribe
function formatearDNI(input){

    let valor = input.value.replace(/\D/g,"");

    if(valor){
        input.value = Number(valor).toLocaleString("es-AR");
    }else{
        input.value = "";
    }

}

//funcion para permitir solo numeros en el campo del dni
function soloNumeros(event){
    if(!/[0-9]/.test(event.key)){
        event.preventDefault();
    }
}

//funcion para formatear la fecha a medida que se escribe, agregando las barras y limitando a 8 numeros
function formatearFechaInput(input){

    let valor = input.value.replace(/\D/g,"");

    if(valor.length > 8){
        valor = valor.substring(0,8);
    }

    let resultado = "";

    if(valor.length >= 1){
        resultado = valor.substring(0,2);
    }

    if(valor.length >= 3){
        resultado += "/" + valor.substring(2,4);
    }

    if(valor.length >= 5){
        resultado += "/" + valor.substring(4);
    }

    input.value = resultado;

}

//funcion para validar la fecha ingresada, corrigiendo el formato y verificando que sea una fecha real
function validarFechaFinal(input){

    let valor = input.value.trim();

    if(!valor) return;

    let partes = valor.split("/");

    if(partes.length !== 3) return;

    let dia = parseInt(partes[0],10);
    let mes = parseInt(partes[1],10);
    let anio = partes[2];

    // completar ceros
    let diaStr = String(dia).padStart(2,"0");
    let mesStr = String(mes).padStart(2,"0");

    // corregir año corto
    if(anio.length === 2){
        if(parseInt(anio) > 30){
            anio = "19" + anio;
        }else{
            anio = "20" + anio;
        }
    }

    anio = parseInt(anio,10);

    // validaciones reales
    if(mes < 1 || mes > 12){
        mostrarMensaje("⚠ Mes inválido", "error");
        input.focus();
        return;
    }

    const diasPorMes = [31, (esBisiesto(anio)?29:28), 31,30,31,30,31,31,30,31,30,31];

    if(dia < 1 || dia > diasPorMes[mes-1]){
        mostrarMensaje("⚠ Día inválido", "error");
        input.focus();
        return;
    }

    input.value = `${diaStr}/${mesStr}/${anio}`;
}

//funcion para determinar si un año es bisiesto (febrero tiene 29 días)
function esBisiesto(anio){
    return (anio % 4 === 0 && anio % 100 !== 0) || (anio % 400 === 0);
}

//FUNCION PARA ELIMINAR COMAS Y CARACTERES EN EL CAMPO nombre_solicitante
function formatearNombreSolicitante(texto){

    return texto
        .replace(/[^\p{L}\s]/gu, "") // 🔥 permite letras con acento
        .replace(/\s+/g, " ")        // evita espacios duplicados
        .trim();

}

function normalizarMayusculas(texto){
    return texto
        .toLocaleUpperCase("es-AR")
        .normalize("NFC");
}