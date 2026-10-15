/**
 * Assistant d'import CSV (comme « Données → À partir d'un fichier texte/CSV » d'Excel) :
 * encodage, séparateur, séparateur décimal, ligne d'en-têtes, avec un aperçu du résultat.
 * Les réglages sont détectés automatiquement ; il suffit en général de cliquer sur « Charger ».
 */
import { t } from '../i18n';
import { csvToRows, decodeCsvBytes, detectCsvFormat, parseCsvValue, type CsvEncoding, type CsvImportOptions } from '../io/csv';
import { box, checkbox, field, openDialog, select, textInput } from '../ui/dialog';

const PREVIEW_ROWS = 12;
const PREVIEW_COLUMNS = 12;

export interface CsvImportChoice {
  text: string;
  options: CsvImportOptions;
}

/** Ouvre l'assistant. Renvoie le texte décodé et les options, ou null si l'utilisateur annule. */
export async function csvImportDialog(fileName: string, bytes: Uint8Array): Promise<CsvImportChoice | null> {
  const detected = decodeCsvBytes(bytes);
  const format = detectCsvFormat(detected.text);
  const knownDelimiters = [';', ',', '\t', '|', ' '];

  const encoding = select([['utf-8', t('csv.utf8')], ['windows-1252', t('csv.windows1252')]], detected.encoding);
  const delimiter = select(
    [[';', t('csv.semicolon')], [',', t('csv.comma')], ['\t', t('csv.tab')], ['|', t('csv.pipe')], [' ', t('csv.space')], ['other', t('csv.other')]],
    knownDelimiters.includes(format.delimiter) ? format.delimiter : 'other',
  );
  const otherDelimiter = textInput(knownDelimiters.includes(format.delimiter) ? '' : format.delimiter);
  otherDelimiter.maxLength = 1;
  otherDelimiter.setAttribute('aria-label', t('csv.otherLabel'));
  const decimal = select([[',', t('csv.decimalComma')], ['.', t('csv.decimalDot')]], format.decimal);
  const header = checkbox(t('csv.header'), looksLikeHeader(detected.text, format.delimiter, format.decimal));
  const preview = document.createElement('div');
  preview.className = 'g-csv-preview';
  preview.dataset.testid = 'csv-preview';

  const delimiterRow = document.createElement('div');
  delimiterRow.className = 'g-row';
  delimiterRow.append(delimiter, otherDelimiter);
  const settings = document.createElement('div');
  settings.className = 'g-csv-settings';
  settings.append(field(t('csv.encoding'), encoding), field(t('csv.delimiter'), delimiterRow), field(t('csv.decimal'), decimal));

  const current = () => {
    const text = decodeCsvBytes(bytes, encoding.value as CsvEncoding).text;
    const sep = delimiter.value === 'other' ? otherDelimiter.value : delimiter.value;
    return { text, options: { delimiter: sep || undefined, decimal: decimal.value as '.' | ',', header: header.input.checked } };
  };
  const render = () => {
    otherDelimiter.hidden = delimiter.value !== 'other';
    const { text, options } = current();
    renderPreview(preview, text, options);
  };
  for (const el of [encoding, delimiter, decimal, header.input]) el.addEventListener('change', render);
  otherDelimiter.addEventListener('input', render);
  render();

  let choice: CsvImportChoice | null = null;
  const ok = await openDialog({
    title: t('csv.title', { name: fileName }),
    body: box(settings, header.el, preview),
    okLabel: t('csv.load'),
    testId: 'dialog-csv',
    wide: true,
    onOk: () => {
      choice = current();
      return true;
    },
  });
  return ok ? choice : null;
}

/** En-têtes probables : la première ligne n'est que du texte et la suivante contient des nombres ou des dates. */
function looksLikeHeader(text: string, delimiter: string, decimal: '.' | ','): boolean {
  const rows = csvToRows(text.slice(0, 20000), delimiter).rows;
  if (rows.length < 2) return false;
  const kinds = (row: string[]) => row.map((raw) => parseCsvValue(raw, decimal)).filter(Boolean);
  const first = kinds(rows[0]);
  const second = kinds(rows[1]);
  return first.length > 0 && first.every((c) => typeof c!.v === 'string') && second.some((c) => typeof c!.v === 'number');
}

function renderPreview(container: HTMLElement, text: string, options: CsvImportOptions): void {
  // L'aperçu ne lit que le début du fichier.
  const rows = csvToRows(text.slice(0, 50000), options.delimiter).rows.slice(0, PREVIEW_ROWS);
  const table = document.createElement('table');
  rows.forEach((row, r) => {
    const tr = document.createElement('tr');
    for (const raw of row.slice(0, PREVIEW_COLUMNS)) {
      const td = document.createElement(r === 0 && options.header ? 'th' : 'td');
      const cell = parseCsvValue(raw, options.decimal ?? ',');
      td.textContent = raw.trim();
      if (typeof cell?.v === 'number') td.className = 'num';
      tr.append(td);
    }
    table.append(tr);
  });
  const columns = Math.max(0, ...rows.map((r) => r.length));
  const info = document.createElement('p');
  info.className = 'g-muted';
  info.textContent = t('csv.previewInfo', { columns: String(columns) });
  container.replaceChildren(info, table);
}
