const xmlFields = ['GLOBALID', 'CLASSID', 'OKTMO', 'ZNP_TYPE', 'MXDOS', 'TR_DOS', 'RZUDOS', 'POPULATIONF', 'POPULATIONP'];
const integerFields = ['ZNP_TYPE', 'MXDOS', 'TR_DOS', 'RZUDOS'];
const uuid = '[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}';
const uuidPattern = new RegExp(`^(?:urn:uuid:)?${uuid}$|^\\{${uuid}\\}$`);

function sheetRows(sheet) {
    const rows = new Map();
    for (const [address, cell] of Object.entries(sheet)) {
        const match = /^([A-Z]+)([1-9]\d*)$/.exec(address);
        if (!match) continue;
        const [, column, number] = match;
        const index = Number(number);
        if (!rows.has(index)) rows.set(index, new Map());
        rows.get(index).set(column, cell);
    }
    return rows;
}

function readColumns(header) {
    const columns = new Map();
    const worksheetLabels = new Map([
        ['Ед. пешком', 'unit'], ['Источник в НГП', 'source'], ['Условия / примечание', 'note'],
    ]);
    for (const [column, cell] of header) {
        const name = textValue(cell.v);
        const parts = name.split(/\r?\n/).map(part => part.trim());
        const key = parts.find(part => xmlFields.includes(part)) ?? worksheetLabels.get(parts[0]);
        if (!key) continue;
        if (columns.has(key)) throw new Error(`В строке заголовков повторяется колонка «${parts[0]}».`);
        columns.set(key, column);
    }
    const missing = [...xmlFields, 'unit'].filter(key => !columns.has(key));
    if (missing.length) {
        const names = missing.map(key => key === 'unit' ? 'Ед. пешком' : key);
        throw new Error(`В строке 4 листа «Заполнение» отсутствуют колонки: ${names.join(', ')}. Скопируйте данные в исходный шаблон.`);
    }
    return columns;
}

function textValue(value) {
    return typeof value === 'string' || typeof value === 'number' ? String(value).trim() : '';
}

function canonicalUuid(value) {
    return value.toLowerCase().replace(/^urn:uuid:|[{}]/g, '');
}

export function convertWorkbook(workbook, schema, createId) {
    const sheet = workbook.Sheets['Заполнение'];
    if (!sheet) throw new Error('В файле нет листа «Заполнение». Используйте шаблон ФГИС ТП.');
    const rows = sheetRows(sheet);
    const columns = readColumns(rows.get(4) ?? new Map());
    const errors = [];
    const records = [];
    const seenKeys = new Map();
    const seenIds = new Map();

    for (const number of [...rows.keys()].sort((a, b) => a - b)) {
        if (number <= 4) continue;
        const row = rows.get(number);
        const cells = [...columns].map(([key, column]) => [key, row.get(column)]);
        const values = Object.fromEntries(cells.map(([key, cell]) => [key, textValue(cell?.t === 'z' ? undefined : cell?.v)]));
        const invalid = cells.filter(([key, cell]) => {
            if (!cell || key === 'source' || key === 'note') return false;
            return cell.t === 'e' || (cell.f && (cell.t === 'z' || cell.v == null)) ||
                (cell.v != null && typeof cell.v !== 'string' && typeof cell.v !== 'number');
        });
        if (!Object.values(values).some(Boolean) && !invalid.length) continue;

        const issues = invalid.map(([key]) => `${key} — введите значение или сохраните пересчитанную формулу`);
        if (!/^\d{8}$|^\d{11}$/.test(values.OKTMO)) issues.push('ОКТМО должен содержать 8 или 11 цифр');
        const className = schema.classes.get(values.CLASSID);
        if (!className) issues.push('выберите CLASSID из справочника XSD');
        for (const key of integerFields) {
            const value = values[key];
            if (!value && key !== 'ZNP_TYPE') continue;
            const minimum = key === 'ZNP_TYPE' ? 0 : 1;
            if (!/^\d+$/.test(value) || !Number.isSafeInteger(Number(value)) || Number(value) < minimum) {
                issues.push(`${key} — укажите целое число от ${minimum}`);
            } else {
                values[key] = String(Number(value));
            }
        }
        if (values.MXDOS && !['м', 'мин'].includes(values.unit)) issues.push('выберите «м» или «мин» для MXDOS');
        if (!values.MXDOS && values.unit) issues.push('указана единица измерения без MXDOS');
        for (const key of ['POPULATIONF', 'POPULATIONP']) {
            const value = values[key].replace(',', '.');
            if (!value) continue;
            if (!/^\+?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(value) || !Number.isFinite(Number(value))) {
                issues.push(`${key} — укажите неотрицательную численность в людях`);
            } else {
                values[key] = String(Number(value));
            }
        }
        if (values.GLOBALID && !uuidPattern.test(values.GLOBALID)) issues.push('GLOBALID должен быть UUID');
        if (issues.length) {
            errors.push({ row: number, issues });
            continue;
        }

        const pair = `${values.OKTMO}|${values.CLASSID}`;
        if (seenKeys.has(pair)) issues.push(`ОКТМО и CLASSID уже указаны в строке ${seenKeys.get(pair)}`);
        else seenKeys.set(pair, number);
        if (values.GLOBALID) {
            const id = canonicalUuid(values.GLOBALID);
            if (seenIds.has(id)) issues.push(`GLOBALID уже указан в строке ${seenIds.get(id)}`);
            else seenIds.set(id, number);
        }
        if (issues.length) errors.push({ row: number, issues });
        else records.push({ className, values });
    }

    if (errors.length) return { errors };
    if (!records.length) throw new Error('На листе «Заполнение» нет данных. Начните ввод с пятой строки.');

    const lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        `<fgistp:FeatureCollection xmlns:fgistp="${schema.namespace}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="${schema.namespace} fgistp-ngp__11.xsd">`,
    ];
    for (const { className, values } of records) {
        const fields = { ...values, GLOBALID: values.GLOBALID || createId() };
        lines.push('  <fgistp:featureMember>', `    <fgistp:${className}>`);
        if (fields.MXDOS) lines.push(`      <!-- MXDOS: ${fields.unit} -->`);
        for (const key of xmlFields) {
            if (fields[key] !== '') lines.push(`      <fgistp:${key}>${fields[key]}</fgistp:${key}>`);
        }
        lines.push(`    </fgistp:${className}>`, '  </fgistp:featureMember>');
    }
    lines.push('</fgistp:FeatureCollection>', '');
    return { xml: lines.join('\n'), count: records.length };
}
