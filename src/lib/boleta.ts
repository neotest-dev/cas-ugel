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
  totalLiquido: string;
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
          "Ingr.: Termino: ",
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
  lines.push("Leyenda Mensual              :");


  // 🔥 BLOQUE PENSIONES FIJO
  if (isONP) {
    lines.push(
        row2(
            "Reg.Pensionario              : ",
            "ONP /W",
            "CFija: ",
            w.aporteObligatorio
        )
    );
  } else {
    lines.push(
        row2(
            "Reg.Pensionario              : ",
            `${w.sistemaPensionario} / ${w.cussp}`,
            "CFija: ",
            w.aporteObligatorio
        )
    );
  }

  lines.push(
      row2(
          "FAfiliacion                  : ",
          w.fechaAfiliacion,
          "CVariable: ",
          w.comision
      )
  );

  lines.push(
      row2(
          "FDevengue                    : ",
          w.fechaDevengue,
          "Seguro: ",
          w.primaSeguro
      )
  );

  lines.push(
      "------------------------------------------------------------------------"
  );

  lines.push(
      `+Honorario            S/.  ${w.montoMensual}`
  );

  lines.push(
      `-${w.sistemaPensionario.padEnd(20, " ")} S/.  ${descuentoPension}`
  );

  // ESPACIOS FIJOS
  for (let i = 0; i < 7; i++) {
    lines.push("");
  }

  lines.push(
      "------------------------------------------------------------------------"
  );

  lines.push(
      `T-HONORARIO S/.  ${w.montoMensual}   T-DSCTO S/.${
          w.totalDscto}   T-LIQUI S/.  ${w.totalLiquido}`
  );

  lines.push(
      `MImponible  S/.  ${w.montoMensual}`
  );

  lines.push("Mensajes :");

  return lines.join("\n");
}