export interface Worker {
    n: string;
    dni: string;
    apPaterno: string;
    apMaterno: string;
    nombres: string;
    fechaNac: string;
    cargo: string;
    codEssalud: string;
    cuentaBanco: string;
    leyendaRD: string;
    leyendaMensual: string;
    sistemaPensionario: string;
    cussp: string;
    fechaAfiliacion: string;
    fechaDevengue: string;
    aporteObligatorio: string;
    comision: string;
    primaSeguro: string;
    montoMensual: string;
    descuentoPension: string;
    onp: string;
    prima: string;
    integra: string;
    profuturo: string;
    habitat: string;
    totalDscto: string;
    otrosDsctos: string;
    dsctoEntidades: string;
    dsctoJudicial: string;
    totalLiquido: string;
}

export interface CategoriaPlanilla {
    id: string;
    label: string;
    match: string[];
}

export const CATEGORIAS_PLANILLA: CategoriaPlanilla[] = [
    { id: "sede", label: "CAS SEDE", match: ["SEDE"] },
    { id: "jec", label: "CAS JEC", match: ["JEC"] },
    { id: "orquestando", label: "CAS ORQUESTANDO", match: ["ORQUESTANDO"] },
    { id: "seho", label: "CAS SEHO", match: ["SEHO", "HOSPITALARIO"] },
    { id: "ebe", label: "CAS EBE INCLUSIVAS", match: ["EBE", "INCLUSIVAS"] },
    { id: "winanq", label: "CAS WINANQ", match: ["WINANQ", "WINAQ", "WIÑANQ", "WIÑAQ"] },
    { id: "convivencia", label: "CAS CONVIVENCIA", match: ["CONVIVENCIA"] },
    { id: "mantenimiento", label: "CAS MANTENIMIENTO", match: ["MANTENIMIENTO"] },
];

export function normalizeCategoryText(text: string): string {
    return text
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toUpperCase();
}

export function inferCategoryIdFromText(text: string): string | null {
    const normalized = normalizeCategoryText(text);
    const matched = CATEGORIAS_PLANILLA.find((category) =>
        category.match.some((match) => normalized.includes(normalizeCategoryText(match)))
    );

    return matched?.id ?? null;
}

const MESES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
];

export function parsePeriodFromFilename(
    filename: string
): { mes: string; anio: string } {
    const m = filename.match(/(\d{1,2})[\s_-]+(\d{4})/);

    if (m) {
        const mesIdx = Number(m[1]) - 1;

        if (mesIdx >= 0 && mesIdx < 12) {
            return {
                mes: MESES[mesIdx],
                anio: m[2]
            };
        }
    }

    const now = new Date();

    return {
        mes: MESES[now.getMonth()],
        anio: String(now.getFullYear())
    };
}




/* YA NO LEE EXCEL AQUI
   Excel lo procesa excelWorker.ts
*/


export function buildBoletaText(
    w: Worker,
    mes: string,
    anio: string
): string {
    const isONP = w.sistemaPensionario
        .toUpperCase()
        .includes("ONP");

    const apellidos =
        `${w.apPaterno} ${w.apMaterno}`.trim();

    const descuentoPension =
        w.descuentoPension ||
        w.onp ||
        w.prima ||
        w.integra ||
        w.profuturo ||
        w.habitat ||
        "0.00";

    // 🔥 FUNCION PARA ALINEAR SIN QUE SE MUEVA NADA
    const row2 = (
        leftLabel: string,
        leftValue: string,
        rightLabel = "",
        rightValue = ""
    ): string => {
        const left = `${leftLabel}${leftValue ?? ""}`;
        const right = rightLabel
            ? `${rightLabel}${rightValue ?? ""}`
            : "";

        return left.padEnd(52, " ") + right;
    };



    const lines: string[] = [];

    lines.push(`BOLETA N° ${w.n}`);
    lines.push("DIRECCION REGIONAL LA LIBERTAD");
    lines.push("*B9 UGEL 04 SUR ESTE");
    lines.push("RUC - 20539889622");
    lines.push(`${mes} - ${anio}`);
    lines.push("");

    lines.push(row2("Apellidos                    : ", apellidos));
    lines.push(row2("Nombres                      : ", w.nombres));
    lines.push(row2("Fecha de Nacimiento          : ", w.fechaNac));
    lines.push(
        row2(
            "Documento de Identidad       : ",
            `(Lib.Electoral o D.N.) ${w.dni}`
        )
    );
    lines.push(
        row2(
            "Establecimiento              : ",
            "UGEL Nº 04 TRUJILLO SUR ESTE"
        )
    );
    lines.push(row2("Cargo                        : ", w.cargo));
    lines.push(
        row2(
            "Tipo de Servidor             : ",
            "ADMINISTRATIVO CONTRATADO"
        )
    );
    lines.push(
        row2(
            "Regimen Laboral              : ",
            "D.LEG.Nº 1057 - CAS"
        )
    );
    lines.push(
        row2(
            "Niv.Mag./Grupo Ocup./Horas   : ",
            "0/0/40 Horas"
        )
    );
    lines.push(
        row2(
            "Tiempo de Servicio (AA-MM-DD): ",
            `-- ESSALUD : ${w.codEssalud}`,

        )
    );
    lines.push(
        row2(
            "Fecha de Registro            : ",
            `Ingr.: ${w.fechaAfiliacion} Termino: ${w.fechaDevengue}`
        )
    );
    lines.push(
        row2(
            "Cta. TeleAhorro o Nro.Cheque : ",
            `CTA- ${w.cuentaBanco}`
        )
    );
    lines.push(
        row2(
            "Leyenda Permanente           : ",
            w.leyendaRD
        )
    );
    lines.push(
        row2(
            "Leyenda Mensual              : ",
            w.leyendaMensual
        )
    );


    // 🔥 BLOQUE PENSIONES
    if (isONP) {
        lines.push(
            row2(
                "Reg.Pensionario              : ",
                "ONP /W"
            )
        );
    } else {
        lines.push(
            row2(
                "Reg.Pensionario              : ",
                `AFP / ${w.cussp}`
            )
        );
    }

    lines.push(
        row2(
            "FAfiliacion                  : ",
            w.fechaAfiliacion
        )
    );

    lines.push(
        row2(
            "FDevengue                    : ",
            w.fechaDevengue
        )
    );

    lines.push(
        "------------------------------------------------------------------------"
    );

    lines.push(
        `PAGO TOTAL MENSUAL            S/.  ${w.montoMensual}`
    );

    const displayPensionSystemName = isONP ? w.sistemaPensionario : "AFP";

    lines.push(
        `-${displayPensionSystemName.padEnd(28, " ")} S/.  ${descuentoPension}`
    );

    let addedLines = 0;
    if (w.otrosDsctos && Number(w.otrosDsctos) > 0) {
        lines.push(`-OTROS DSCTOS`.padEnd(29, " ") + ` S/.  ${w.otrosDsctos}`);
        addedLines++;
    }
    if (w.dsctoEntidades && Number(w.dsctoEntidades) > 0) {
        lines.push(`-DESCUENTO ENTIDADES`.padEnd(29, " ") + ` S/.  ${w.dsctoEntidades}`);
        addedLines++;
    }
    if (w.dsctoJudicial && Number(w.dsctoJudicial) > 0) {
        lines.push(`-DSCTO JUDICIAL`.padEnd(29, " ") + ` S/.  ${w.dsctoJudicial}`);
        addedLines++;
    }

    // ESPACIOS FIJOS
    const emptySpaces = Math.max(0, 7 - addedLines);
    for (let i = 0; i < emptySpaces; i++) {
        lines.push("");
    }

    lines.push(
        "------------------------------------------------------------------------"
    );

    lines.push(
        `T-DSCTO S/.${w.totalDscto}   T-LIQUI S/.  ${w.totalLiquido}`
    );

    lines.push("Mensajes :");

    return lines.join("\n");
}
