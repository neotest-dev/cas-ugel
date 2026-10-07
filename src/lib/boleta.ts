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
    { id: "seho", label: "CAS HOSPITALARIOS", match: ["HOSPITALARIO", "HOSPITALARIOS", "SEHO"] },
    { id: "ebe", label: "CAS MEDICA-CEBE", match: ["MEDICA-CEBE", "MEDICA CEBE", "MEDICA", "CEBE", "EBE", "INCLUSIVAS"] },
    { id: "winanq", label: "CAS WIÑANQ", match: ["WIÑANQ", "WIÑAQ", "WINANQ", "WINAQ"] },
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

    // Para AFP: suma integra + profuturo + habitat + prima (el que aplique)
    // Para ONP: usa el campo onp directamente
    // Se usa parseFloat para evitar el bug de "0.00" siendo string truthy con ||
    let descuentoPension: string;
    if (isONP) {
        const v = parseFloat(w.onp || "0") || parseFloat(w.descuentoPension || "0");
        descuentoPension = v > 0 ? v.toFixed(2) : "0.00";
    } else {
        const afpSum =
            (parseFloat(w.integra || "0") || 0) +
            (parseFloat(w.profuturo || "0") || 0) +
            (parseFloat(w.habitat || "0") || 0) +
            (parseFloat(w.prima || "0") || 0);
        if (afpSum > 0) {
            descuentoPension = afpSum.toFixed(2);
        } else {
            const fallback = parseFloat(w.descuentoPension || "0") || parseFloat(w.onp || "0");
            descuentoPension = fallback > 0 ? fallback.toFixed(2) : "0.00";
        }
    }

    // FUNCION PARA ALINEAR
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


    // BLOQUE PENSIONES
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

    const formatMoney = (val: string | number | undefined | null) => {
        const num = parseFloat(String(val ?? "0"));
        return isNaN(num) ? "0.00" : num.toFixed(2);
    };

    lines.push(
        `PAGO TOTAL MENSUAL            S/.  ${formatMoney(w.montoMensual)}`
    );

    if (isONP) {
        // ONP: mostrar en una sola línea con el nombre del sistema
        const label = w.sistemaPensionario || "ONP";
        lines.push(`-${label.padEnd(28, " ")} S/.  ${formatMoney(descuentoPension)}`);
    } else {
        // AFP: mostrar cada fondo con valor propio en su línea
        const afpFunds: Array<{ label: string; val: number }> = [
            { label: "AFP INTEGRA",   val: parseFloat(w.integra   || "0") || 0 },
            { label: "AFP PROFUTURO", val: parseFloat(w.profuturo || "0") || 0 },
            { label: "AFP HABITAT",   val: parseFloat(w.habitat   || "0") || 0 },
            { label: "AFP PRIMA",     val: parseFloat(w.prima     || "0") || 0 },
        ].filter(f => f.val > 0);

        if (afpFunds.length > 0) {
            for (const fund of afpFunds) {
                lines.push(`-${fund.label.padEnd(28, " ")} S/.  ${fund.val.toFixed(2)}`);
            }
        } else {
            // Fallback si no se detectó el fondo específico
            lines.push(`-AFP                          S/.  ${formatMoney(descuentoPension)}`);
        }
    }


    let addedLines = 0;
    if (w.otrosDsctos && Number(w.otrosDsctos) > 0) {
        lines.push(`-OTROS DSCTOS`.padEnd(29, " ") + ` S/.  ${formatMoney(w.otrosDsctos)}`);
        addedLines++;
    }
    if (w.dsctoEntidades && Number(w.dsctoEntidades) > 0) {
        lines.push(`-DESCUENTO ENTIDADES`.padEnd(29, " ") + ` S/.  ${formatMoney(w.dsctoEntidades)}`);
        addedLines++;
    }
    if (w.dsctoJudicial && Number(w.dsctoJudicial) > 0) {
        lines.push(`-DSCTO JUDICIAL`.padEnd(29, " ") + ` S/.  ${formatMoney(w.dsctoJudicial)}`);
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
        `T-DSCTO S/.${formatMoney(w.totalDscto)}   T-LIQUI S/.  ${formatMoney(w.totalLiquido)}`
    );

    lines.push("Mensajes :");

    return lines.join("\n");
}
