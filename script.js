document
    .getElementById("btnConvertir")
    .addEventListener("click", convertir);

function extraerCaracteristica(texto, nombre) {

    const regex = new RegExp(
        `${nombre}:\\s*([0-9.,]+)`,
        "i"
    );

    const resultado = texto.match(regex);

    if (!resultado)
        return null;

    return convertirDimension(resultado[1]);
}

function convertirDimension(valor) {

    if (!valor)
        return null;

    valor = valor.replace(",", ".");

    const numero = parseFloat(valor);

    return isNaN(numero)
        ? null
        : numero;
}


function convertirNumero(valor) {

    if (!valor) return null;

    valor = valor
        .replace(/\./g, "")
        .replace(",", ".");

    const numero = parseFloat(valor);

    return isNaN(numero) ? null : numero;
}

async function convertir() {

    const archivos =
        document.getElementById("archivos").files;

    if (archivos.length === 0) {
        alert("Seleccione uno o más TXT");
        return;
    }

    let datos = [];

    for (const archivo of archivos) {

        const contenido = await archivo.text();

        const lineas =
            contenido.split(/\r?\n/);

        for (const line of lineas) {

            if (!/^\s*\d{5,6}\b/.test(line))
                continue;

            const nro_ot =
                line.substring(2, 14).trim();

            const localidad =
                line.substring(14, 39).trim();

            const domicilio =
                line.substring(39, 65).trim();

            const altura =
                line.substring(65, 74).trim();

            const fechaSol =
                line.substring(74, 85).trim();

            const fechaEjec =
                line.substring(85, 96).trim();

            const fechaCie =
                line.substring(96, 107).trim();

            const vereda =
                line.substring(107, 120).trim();

            const tipoOT =
                line.substring(120, 131).trim();

            const prioridad =
                line.substring(131, 137).trim();

            const distancia =
                line.substring(137, 145).trim();

            const resto =
                line.substring(145);

            const precioMatch =
                resto.match(
                    /(\d{1,3}(?:\.\d{3})*,\d{2})/
                );

            const precio =
                precioMatch
                    ? convertirNumero(
                        precioMatch[1]
                    )
                    : null;
            let caracteristicas = resto;

if (precioMatch) {

    caracteristicas =
        resto.substring(
            precioMatch.index +
            precioMatch[0].length
        );
}

const cant =
    extraerCaracteristica(
        caracteristicas,
        "Cnt"
    );

const A =
    extraerCaracteristica(
        caracteristicas,
        "Anc"
    );

const L =
    extraerCaracteristica(
        caracteristicas,
        "Lar"
    );

const P =
    extraerCaracteristica(
        caracteristicas,
        "Pro"
    );

let M3 = "";

if (
    A !== null &&
    L !== null &&
    P !== null
) {
    M3 = Number(
        (A * L * P).toFixed(3)
    );
}
``
            datos.push({

                "N° OT": nro_ot,
                "Localidad": localidad,
                "Domicilio": domicilio,
                "Altura": altura,
                "Fecha Sol.": fechaSol,
                "Fecha Ejec.": fechaEjec,
                "Fecha Cie.": fechaCie,
                "Vereda": vereda,
                "Tipo OT": tipoOT,
                "Prior.": prioridad,
                "Dist.": distancia,
                "Precio": precio,
                "Cant.": cant,
                "A": A,
                "L": L,
                "P": P,
                "M³": M3,
                "Observaciones": ""

            });
        }
    }

    const wb =
        XLSX.utils.book_new();

    const ws =
        XLSX.utils.json_to_sheet(datos);

    XLSX.utils.book_append_sheet(
        wb,
        ws,
        "Correcciones"
    );

    XLSX.writeFile(
        wb,
        "Correcciones_Automaticas.xlsx"
    );

    document.getElementById("resultado")
        .innerHTML =
        `<p>✅ Excel generado con ${datos.length} registros.</p>`;
}
``