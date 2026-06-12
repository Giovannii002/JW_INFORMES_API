import xlsx from 'xlsx';
import fs from 'fs';

const MESES = [
    { mes: 'Noviembre 2025', archivo: './Noviembre.xls' },
    { mes: 'Diciembre 2025', archivo: './Diciembre.xls' },
    { mes: 'Enero 2026', archivo: './Enero.xls' },
    { mes: 'Febrero 2026', archivo: './Febrero.xls' },
    { mes: 'Marzo 2026', archivo: './Marzo.xls' },
    { mes: 'Abril 2026', archivo: './Abril.xls' },
];

const normalizeText = (value) => {
    if (value === null || value === undefined) return '';
    return String(value).trim().toUpperCase();
};

const toNumber = (value) => {
    if (value === null || value === undefined || value === '') return 0;

    // por si vienen como texto, con comas o prefijos de exportacion tipo "-->1"
    const cleaned = String(value).replace(/,/g, '').replace(/-->/g, '').trim();
    const num = Number(cleaned);

    return Number.isNaN(num) ? 0 : num;
};

const getValue = (row, ...keys) => {
    for (const key of keys) {
        if (row[key] !== undefined) return row[key];
        if (row[`-->${key}`] !== undefined) return row[`-->${key}`];

        if (normalizeText(key).includes('PARTICIPO MINISTERIO')) {
            const matchingKey = Object.keys(row).find((columnName) => (
                normalizeText(columnName).includes('PARTICIPO MINISTERIO')
            ));

            if (matchingKey) return row[matchingKey];
        }
    }

    return null;
};

const leerFilas = (archivo) => {
    const workbook = xlsx.readFile(archivo);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(sheet, { defval: null });
};

const calcularResumenMensual = (rows, mes) => {
    let totalHoras = 0;
    let totalEstudios = 0;
    let horasPrecursoresRegulares = 0;
    const precursoresRegularesConHora = new Set();

    for (const row of rows) {
        const nombre = String(getValue(row, 'Nombre') ?? '').trim();
        const privilegio = normalizeText(getValue(row, 'Privilegio'));
        const horas = toNumber(getValue(row, 'Horas'));
        const estudios = toNumber(getValue(row, 'Estudios'));

        totalHoras += horas;
        totalEstudios += estudios;

        if (privilegio === 'PREC. REGULAR') {
            horasPrecursoresRegulares += horas;

            if (nombre && horas > 0) {
                precursoresRegularesConHora.add(nombre);
            }
        }
    }

    const cantidadPrecursoresRegulares = precursoresRegularesConHora.size;
    const promedioHorasPrecursorRegular = cantidadPrecursoresRegulares === 0
        ? 0
        : horasPrecursoresRegulares / cantidadPrecursoresRegulares;

    return {
        mes,
        totalHoras,
        totalEstudios,
        horasPrecursoresRegulares,
        cantidadPrecursoresRegulares,
        promedioHorasPrecursorRegular: Number(promedioHorasPrecursorRegular.toFixed(2)),
    };
};

const padCell = (value, width, align = 'left') => {
    const text = String(value);
    return align === 'right' ? text.padStart(width, ' ') : text.padEnd(width, ' ');
};

const generarResumenTxtLegible = (resumen, filePath = './resumen_mensual.txt') => {
    const headers = [
        'Mes',
        'Total horas',
        'Total estudios',
        'Horas prec. regulares',
        'Cant. prec. regulares',
        'Promedio hrs/prec. regular',
    ];

    const rows = resumen.map((item) => [
        item.mes,
        item.totalHoras,
        item.totalEstudios,
        item.horasPrecursoresRegulares,
        item.cantidadPrecursoresRegulares,
        item.promedioHorasPrecursorRegular,
    ]);

    const widths = headers.map((header, index) => Math.max(
        header.length,
        ...rows.map((row) => String(row[index]).length),
    ));

    const separator = `+-${widths.map((width) => '-'.repeat(width)).join('-+-')}-+`;

    const headerLine = `| ${headers
        .map((header, index) => padCell(header, widths[index]))
        .join(' | ')} |`;

    const bodyLines = rows.map((row) => (`| ${row
        .map((cell, index) => padCell(cell, widths[index], index === 0 ? 'left' : 'right'))
        .join(' | ')} |`));

    const totalHoras = resumen.reduce((acc, item) => acc + item.totalHoras, 0);
    const totalEstudios = resumen.reduce((acc, item) => acc + item.totalEstudios, 0);

    const lines = [
        'RESUMEN MENSUAL',
        '--------------',
        '',
        separator,
        headerLine,
        separator,
        ...bodyLines,
        separator,
        '',
        `Total general de horas: ${totalHoras}`,
        `Total general de estudios: ${totalEstudios}`,
    ];

    fs.writeFileSync(filePath, `${lines.join('\n')}\n`, 'utf8');
};

const resumenMensual = [];

for (const mesConfig of MESES) {
    if (!fs.existsSync(mesConfig.archivo)) {
        console.warn(`No se encontro el archivo de ${mesConfig.mes}: ${mesConfig.archivo}`);
        continue;
    }

    const rows = leerFilas(mesConfig.archivo);
    resumenMensual.push(calcularResumenMensual(rows, mesConfig.mes));
}

if (resumenMensual.length === 0) {
    console.error('No se encontro ningun archivo mensual para procesar.');
    process.exit(1);
}

generarResumenTxtLegible(resumenMensual);
fs.writeFileSync('./resumen_mensual.json', `${JSON.stringify(resumenMensual, null, 2)}\n`, 'utf8');

console.table(resumenMensual);
console.log('Resumen generado en ./resumen_mensual.txt y ./resumen_mensual.json');
