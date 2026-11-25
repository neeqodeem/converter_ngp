import { convertWorkbook } from './converter.mjs';
import { validateXML } from './assets/vendor/index-browser.mjs';

const form = document.querySelector('#converter');
const input = document.querySelector('#workbook');
const submit = document.querySelector('#convert');
const result = document.querySelector('#result');
const message = document.querySelector('#message');
const download = document.querySelector('#download');
const list = document.querySelector('#errors');
let assetUrl;

function reset() {
    if (assetUrl) URL.revokeObjectURL(assetUrl);
    assetUrl = undefined;
    download.removeAttribute('href');
    download.hidden = true;
    list.replaceChildren();
    list.hidden = true;
    result.hidden = true;
}

function status(state, text) {
    result.dataset.state = state;
    message.textContent = text;
    result.hidden = false;
}

async function loadSchema() {
    const response = await fetch('./assets/fgistp-ngp__11.xsd');
    if (!response.ok) throw new Error('Не удалось загрузить XSD. Обновите страницу и попробуйте ещё раз.');
    const contents = await response.text();
    const tree = new DOMParser().parseFromString(contents, 'application/xml');
    const root = tree.documentElement;
    const namespace = root.getAttribute('targetNamespace');
    const classes = new Map();
    for (const type of root.getElementsByTagNameNS('http://www.w3.org/2001/XMLSchema', 'simpleType')) {
        const name = type.getAttribute('name');
        if (!name.endsWith('_CLASSID')) continue;
        for (const code of type.getElementsByTagNameNS('http://www.w3.org/2001/XMLSchema', 'enumeration')) {
            classes.set(code.getAttribute('value'), name.slice(0, -8));
        }
    }
    if (!namespace || !classes.size) throw new Error('Файл XSD не удалось прочитать.');
    return { contents, namespace, classes };
}

input.addEventListener('change', reset);
form.addEventListener('submit', async event => {
    event.preventDefault();
    reset();
    const file = input.files[0];
    if (!file || !/\.xlsx$/i.test(file.name)) {
        status('error', 'Выберите файл Excel с расширением .xlsx.');
        return;
    }
    submit.disabled = input.disabled = true;
    status('working', 'Читаем таблицу и проверяем данные…');
    try {
        if (!globalThis.XLSX) throw new Error('Не удалось загрузить обработчик Excel. Обновите страницу.');
        let workbook;
        try {
            workbook = globalThis.XLSX.read(await file.arrayBuffer(), { sheetStubs: true });
        } catch {
            throw new Error('Файл Excel не удалось открыть. Убедитесь, что он сохранён в формате .xlsx и доступен без пароля.');
        }
        const schema = await loadSchema();
        const converted = convertWorkbook(workbook, schema, () => crypto.randomUUID());
        if (converted.errors) {
            status('error', `Строк с ошибками: ${converted.errors.length}. Исправьте таблицу и выберите файл снова.`);
            for (const { row, issues } of converted.errors) {
                const item = document.createElement('li');
                item.textContent = `Строка ${row}: ${issues.join('; ')}.`;
                list.append(item);
            }
            list.hidden = false;
            return;
        }
        const checked = await validateXML({ xml: [converted.xml], schema: [schema.contents] });
        if (!checked.valid) {
            console.error(checked.errors);
            throw new Error('XML не прошёл проверку по XSD. Выгрузка не создана.');
        }
        assetUrl = URL.createObjectURL(new Blob([converted.xml], { type: 'application/xml;charset=utf-8' }));
        download.href = assetUrl;
        download.download = file.name.replace(/\.xlsx$/i, '') + '.xml';
        download.hidden = false;
        status('success', `XML прошёл проверку по XSD. Записей: ${converted.count}.`);
    } catch (error) {
        console.error(error);
        status('error', error?.message ?? 'Не удалось выполнить конвертацию.');
    } finally {
        input.disabled = submit.disabled = false;
    }
});
