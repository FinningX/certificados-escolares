function descargarPDF(){
window.print();
}

/*codigo para leer el Excel*/ 
function procesarExcel(){

        const file = document.getElementById("excelFile").files[0];

        const reader = new FileReader();

        reader.onload = function(e){

        const data = new Uint8Array(e.target.result);

        const workbook = XLSX.read(data,{type:'array'});

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const alumnos = XLSX.utils.sheet_to_json(sheet);

        /*para los alertas*/
        const columnasRequeridas = [
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
        const columnasExcel = Object.keys(alumnos[0] || {});

        const columnasFaltantes = columnasRequeridas.filter(col => !columnasExcel.includes(col));

        if(columnasFaltantes.length > 0){

            alert("El archivo Excel tiene columnas faltantes:\n\n" + columnasFaltantes.join("\n"));

            return;

        }
        /*--------------------------------__*/

        generarCertificados(alumnos);

    };

    reader.readAsArrayBuffer(file);

}

function generarPDFMultiple(){

        const elemento = document.getElementById("contenedorPDF");

        const opciones = {

        margin:0,

        filename:"certificados_alumnos.pdf",

        image:{type:'jpeg',quality:1},

        html2canvas:{scale:3},

        jsPDF:{unit:'mm',format:'a4',orientation:'portrait'}

    };

    html2pdf().set(opciones).from(elemento).save();

}

function convertirFechaExcel(fecha){

    if(typeof fecha === "number"){

        const fechaJS = new Date((fecha - 25569) * 86400 * 1000);

        const dia = String(fechaJS.getDate()).padStart(2,'0');
        const mes = String(fechaJS.getMonth() + 1).padStart(2,'0');
        const anio = fechaJS.getFullYear();

        return `${dia}/${mes}/${anio}`;
    }
    
    return fecha;
}

/*funcion que genera los certificados automaticamente*/ 
async function generarCertificados(alumnos){

    const contenedor = document.getElementById("contenedorPDF");

    contenedor.innerHTML = "";

    const modelo = document.querySelector(".certificado");

    for(let alumno of alumnos){

        const copia = modelo.cloneNode(true);
        aplicarGenero(copia, alumno.genero);

        copia.querySelector(".campoAlumno").innerText = alumno.nombre;
        /*aplica separador de miles al dni*/
        copia.querySelector("#dni").value = alumno.dni.toLocaleString("es-AR");

        /*Convierte la fecha a formato 00/00/0000*/
        copia.querySelector("#nacimiento").value = convertirFechaExcel(alumno.nacimiento);

        copia.querySelector("#edad").value = alumno.edad;
        
        copia.querySelector("#grado").value = alumno.grado;

        copia.querySelector("#solicitante").value = alumno.solicitante;

        copia.querySelector("#nombre_solicitante").value = alumno.nombre_solicitante;

        /*aplica separador de miles al dni*/
        copia.querySelector("#dni_solicitante").value = alumno.dni_solicitante.toLocaleString("es-AR");

        copia.querySelector(".campoSolicitante").innerText = alumno.ante;

        copia.querySelector(".campoSolicitante2").innerText = alumno.ante2 || "";

        const linea2 = copia.querySelector(".campoSolicitante2").closest("p");

        if(!alumno.ante2 || alumno.ante2.trim() === ""){
            linea2.style.display = "none";
        }else{
            linea2.style.display = "flex";
        }

        copia.querySelector("#localidad").value = alumno.localidad?.trim();
        copia.style.pageBreakAfter = "auto";

        contenedor.appendChild(copia);

    }

    /*generarPDFMultiple();*/

}

/*Funcion que genera el pdf*/
function generarPDF(nombreAlumno){

        const elemento = document.querySelector(".certificado");

        const opciones = {

        margin:0,

        filename:`Certificado_${nombreAlumno}.pdf`,

        image:{type:'jpeg',quality:1},

        html2canvas:{scale:3},

        jsPDF:{unit:'mm',format:'a4',orientation:'portrait'}

    };

    return html2pdf().set(opciones).from(elemento).save();

}

/*Funcion para detectar el genero y tachar Dn o Dña*/
function aplicarGenero(certificado, genero){

    const dn = certificado.querySelector("#dn");
    const dna = certificado.querySelector("#dna");

    dn.style.textDecoration = "none";
    dna.style.textDecoration = "none";

    if(!genero) return;

    genero = genero.toString().trim().toUpperCase();

    if(genero === "F"){
        dn.style.textDecoration = "line-through";
    }

    if(genero === "M"){
        dna.style.textDecoration = "line-through";
    } 

}

/*funcion para agregar certificado manualmente si falto alguno*/
function agregarCertificadoManual(){

const contenedor = document.getElementById("contenedorPDF");

const plantilla = document.querySelector(".plantilla");

const copia = plantilla.cloneNode(true);

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

copia.querySelectorAll("[contenteditable]").forEach(el=>{
    el.contentEditable = "false";
});

copia.querySelectorAll("input").forEach(el=>{
    el.readOnly = true;
});

contenedor.appendChild(copia);

/* limpiar plantilla */

plantilla.querySelectorAll("input").forEach(el=> el.value="");

plantilla.querySelectorAll("[contenteditable]").forEach(el=> el.innerText="");

}