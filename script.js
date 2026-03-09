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

        return fechaJS.toLocaleDateString("es-AR");

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

        copia.querySelector(".campoAlumno").innerText = alumno.nombre;

        copia.querySelector("#dni").value = alumno.dni;

        copia.querySelector("#nacimiento").value = convertirFechaExcel(alumno.nacimiento);

        copia.querySelector("#edad").value = alumno.edad;

        copia.querySelector("#grado").value = alumno.grado;

        copia.querySelector("#solicitante").value = alumno.solicitante;

        copia.querySelector("#nombre_solicitante").value = alumno.nombre_solicitante;

        copia.querySelector("#dni_solicitante").value = alumno.dni_solicitante;

        copia.querySelector(".campoSolicitante").innerText = alumno.ante;

        copia.querySelector("#localidad").value = alumno.localidad;

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