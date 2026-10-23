/**
 * Graphiques d'Excel 2016 au format « chartEx » (xl/charts/chartExN.xml) : cascade, entonnoir,
 * histogramme et boîte à moustaches. Le dessin les référence dans un mc:AlternateContent, avec un
 * rectangle de remplacement pour les logiciels qui ne les connaissent pas.
 */
import type { ChartSpec, SeriesSpec, StatChartType } from '../../charts/spec';
import { attrs, escapeXml, parseXml } from './ooxml';

export const CT_CHARTEX = 'application/vnd.ms-office.chartex+xml';
export const REL_CHARTEX = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
const NS_CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex';

const LAYOUT: Record<StatChartType, string> = { waterfall: 'waterfall', funnel: 'funnel', histogram: 'clusteredColumn', boxWhisker: 'boxWhisker' };

type Read = (ref: string | undefined) => (string | number | null)[];

const f = (ref: string) => `<cx:f>${escapeXml(ref.replace(/^=/, ''))}</cx:f>`;

function strDim(ref: string, values: (string | number | null)[]): string {
  const pts = values.map((v, i) => (v === null ? '' : `<cx:pt idx="${i}">${escapeXml(String(v))}</cx:pt>`)).join('');
  return `<cx:strDim type="cat">${f(ref)}<cx:lvl ptCount="${values.length}">${pts}</cx:lvl></cx:strDim>`;
}

function numDim(ref: string, values: (string | number | null)[]): string {
  const pts = values.map((v, i) => (typeof v === 'number' ? `<cx:pt idx="${i}">${v}</cx:pt>` : '')).join('');
  return `<cx:numDim type="val">${f(ref)}<cx:lvl ptCount="${values.length}" formatCode="General">${pts}</cx:lvl></cx:numDim>`;
}

const text = (value: string, ref?: string) => `<cx:tx><cx:txData>${ref ? f(ref) : ''}<cx:v>${escapeXml(value)}</cx:v></cx:txData></cx:tx>`;
const fill = (color?: string) => (color ? `<cx:spPr><a:solidFill><a:srgbClr val="${color.replace('#', '').toUpperCase()}"/></a:solidFill></cx:spPr>` : '');
const uid = (i: number) => `{${String(i + 1).padStart(8, '0')}-0001-0000-0000-000000000000}`;

/** ChartSpec (cascade, entonnoir, histogramme, boîte à moustaches) → partie chartEx. */
export function buildChartExXml(spec: ChartSpec & { type: StatChartType }, read: Read): string {
  const type = spec.type;
  const series = type === 'waterfall' || type === 'funnel' ? spec.series.slice(0, 1) : spec.series;
  const data = series
    .map((s, i) => {
      const cat = s.categoriesRef && type !== 'histogram' ? strDim(s.categoriesRef, read(s.categoriesRef)) : '';
      return `<cx:data id="${i}">${cat}${numDim(s.valuesRef, read(s.valuesRef))}</cx:data>`;
    })
    .join('');
  const labels = spec.dataLabels?.show ?? (type === 'waterfall' || type === 'funnel');
  const dataLabels = labels ? `<cx:dataLabels pos="${type === 'funnel' ? 'ctr' : 'outEnd'}"><cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels>` : '';
  const layoutPr = (): string => {
    if (type === 'waterfall') {
      const totals = spec.waterfallTotals ?? [];
      return totals.length ? `<cx:layoutPr><cx:subtotals>${totals.map((i) => `<cx:idx val="${i}"/>`).join('')}</cx:subtotals></cx:layoutPr>` : '';
    }
    if (type === 'histogram') {
      const b = spec.bins ?? {};
      const size = b.count ? `<cx:binCount val="${Math.round(b.count)}"/>` : b.width ? `<cx:binSize val="${b.width}"/>` : '';
      return `<cx:layoutPr><cx:binning${attrs({ intervalClosed: 'r', underflow: b.underflow, overflow: b.overflow })}>${size}</cx:binning></cx:layoutPr>`;
    }
    if (type === 'boxWhisker') {
      return `<cx:layoutPr><cx:visibility meanLine="0" meanMarker="${spec.boxShowMean === false ? 0 : 1}" nonoutliers="0" outliers="${spec.boxShowOutliers === false ? 0 : 1}"/><cx:statistics quartileMethod="exclusive"/></cx:layoutPr>`;
    }
    return '';
  };
  const seriesXml = series
    .map((s: SeriesSpec, i) => {
      const name = s.name ?? (s.nameRef ? String(read(s.nameRef)[0] ?? '') : `Série ${i + 1}`);
      const tx = text(name, s.nameRef);
      return `<cx:series${attrs({ layoutId: LAYOUT[type], uniqueId: uid(i) })}>${tx}${fill(s.color)}${dataLabels}<cx:dataId val="${i}"/>${layoutPr()}</cx:series>`;
    })
    .join('');
  const gap = type === 'funnel' ? 0.06 : type === 'histogram' ? (spec.gapWidth ?? 0) / 100 : type === 'boxWhisker' ? 1 : (spec.gapWidth ?? 50) / 100;
  const axes =
    type === 'funnel'
      ? `<cx:axis id="0"><cx:catScaling gapWidth="${gap}"/><cx:tickLabels/></cx:axis>`
      : `<cx:axis id="0"><cx:catScaling gapWidth="${gap}"/>${spec.xAxisTitle ? `<cx:title>${text(spec.xAxisTitle)}</cx:title>` : ''}<cx:tickLabels/></cx:axis>` +
        `<cx:axis id="1"><cx:valScaling${attrs({ min: spec.axes?.y?.min, max: spec.axes?.y?.max })}/>${spec.yAxisTitle ? `<cx:title>${text(spec.yAxisTitle)}</cx:title>` : ''}<cx:majorGridlines/><cx:tickLabels/></cx:axis>`;
  const title = spec.title ? `<cx:title pos="t" align="ctr" overlay="0">${text(spec.title)}</cx:title>` : '';
  const legendPos = ({ top: 't', bottom: 'b', left: 'l', right: 'r' } as Record<string, string>)[spec.legend ?? 'bottom'];
  const legend = spec.legend === 'none' || type === 'funnel' || type === 'histogram' ? '' : `<cx:legend pos="${legendPos ?? 'b'}" align="ctr" overlay="0"/>`;
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="${NS_CX}">` +
    `<cx:chartData>${data}</cx:chartData>` +
    `<cx:chart>${title}<cx:plotArea><cx:plotAreaRegion>${seriesXml}</cx:plotAreaRegion>${axes}</cx:plotArea>${legend}</cx:chart>` +
    '</cx:chartSpace>'
  );
}

/** Contenu d'une ancre de dessin pour un graphique chartEx (avec remplacement pour les anciens logiciels). */
// Tous les types, entonnoir compris, sous « cx1 » : avec cx2 ou cx4, Excel pour Mac n'affiche que le remplacement (vérifié).
export function chartExFrameXml(_type: StatChartType, id: number, name: string, rid: string): string {
  const [prefix, ns] = ['cx1', 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex'];
  return (
    `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:${prefix}="${ns}" Requires="${prefix}">` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr${attrs({ id, name })}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="${NS_CX}" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>` +
    `<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr${attrs({ id, name })}/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
    '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>' +
    '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="fr-FR" sz="1100"/><a:t>Ce graphique n’est pas disponible dans votre version d’Excel.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent>'
  );
}

type Xml = Record<string, unknown>;
const arr = <T>(v: T | T[] | undefined): T[] => (v === undefined ? [] : Array.isArray(v) ? v : [v]);
const str = (v: unknown): string | undefined => (v === undefined || v === null ? undefined : typeof v === 'object' ? ((v as Xml)['#text'] as string | undefined) : String(v));
const num = (v: unknown): number | undefined => {
  const n = Number(v);
  return v === undefined || v === '' || v === 'auto' || !Number.isFinite(n) ? undefined : n;
};

/** Partie chartEx → ChartSpec (null si le type n'est pas géré : carte, rayons de soleil…). */
export function parseChartExXml(xml: string): ChartSpec | null {
  const space = parseXml<{ chartSpace?: Xml }>(xml).chartSpace;
  const chart = space?.chart as Xml | undefined;
  const region = (chart?.plotArea as Xml | undefined)?.plotAreaRegion as Xml | undefined;
  const all = arr(region?.series as Xml | Xml[]);
  const layout = String(all[0]?.layoutId ?? '');
  const type = (Object.entries(LAYOUT).find(([, l]) => l === layout)?.[0] ?? null) as StatChartType | null;
  if (!type) return null;
  const data = new Map(arr((space?.chartData as Xml | undefined)?.data as Xml | Xml[]).map((d) => [String(d.id), d]));
  const series: SeriesSpec[] = [];
  for (const s of all) {
    const d = data.get(String((s.dataId as Xml | undefined)?.val ?? '0'));
    if (!d) continue;
    const valueRef = str((arr(d.numDim as Xml | Xml[])[0] as Xml | undefined)?.f);
    if (!valueRef) continue;
    const catRef = str((arr(d.strDim as Xml | Xml[])[0] as Xml | undefined)?.f);
    const txData = (s.tx as Xml | undefined)?.txData as Xml | undefined;
    const color = ((((s.spPr as Xml | undefined)?.solidFill as Xml | undefined)?.srgbClr as Xml | undefined)?.val as string | undefined);
    series.push({
      valuesRef: valueRef,
      categoriesRef: catRef,
      nameRef: str(txData?.f),
      name: txData?.f ? undefined : str(txData?.v),
      color: color ? `#${color.toUpperCase()}` : undefined,
    });
  }
  if (!series.length) return null;
  const spec: ChartSpec = { version: 1, type, series };
  const titleV = str((((chart?.title as Xml | undefined)?.tx as Xml | undefined)?.txData as Xml | undefined)?.v);
  if (titleV) spec.title = titleV;
  if (!chart?.legend && type !== 'funnel' && type !== 'histogram') spec.legend = 'none';
  const pos = { t: 'top', l: 'left', r: 'right' }[String((chart?.legend as Xml | undefined)?.pos ?? 'b') as 't' | 'l' | 'r'] as ChartSpec['legend'];
  if (chart?.legend && pos) spec.legend = pos;
  const lp = (all[0]?.layoutPr ?? {}) as Xml;
  if (type === 'waterfall') {
    const idx = arr(((lp.subtotals as Xml | undefined)?.idx as Xml | Xml[]) ?? []).map((x) => Number(x.val)).filter(Number.isInteger);
    if (idx.length) spec.waterfallTotals = idx;
  }
  if (type === 'histogram') {
    const b = (lp.binning ?? {}) as Xml;
    const bins = { width: num((b.binSize as Xml | undefined)?.val), count: num((b.binCount as Xml | undefined)?.val), underflow: num(b.underflow), overflow: num(b.overflow) };
    if (Object.values(bins).some((v) => v !== undefined)) spec.bins = Object.fromEntries(Object.entries(bins).filter(([, v]) => v !== undefined));
  }
  if (type === 'boxWhisker') {
    const v = (lp.visibility ?? {}) as Xml;
    if (String(v.meanMarker) === '0') spec.boxShowMean = false;
    if (String(v.outliers) === '0') spec.boxShowOutliers = false;
  }
  if (all[0]?.dataLabels === undefined && (type === 'waterfall' || type === 'funnel')) spec.dataLabels = { show: false };
  return spec;
}

// --- Style et palette du graphique : obligatoires pour Excel ---
// Sans ces deux parties reliées à chaque graphique chartEx, Excel « répare » le fichier en supprimant tout le dessin
// (vérifié avec Excel pour Mac : « Partie supprimée : /xl/drawings/drawing1.xml »). Les graphiques classiques s'en passent.

export const CT_CHART_STYLE = 'application/vnd.ms-office.chartstyle+xml';
export const CT_CHART_COLORS = 'application/vnd.ms-office.chartcolorstyle+xml';
export const REL_CHART_STYLE = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle';
export const REL_CHART_COLORS = 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';

const NS_CS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** Texte gris (comme Excel : 65 % de la couleur du texte) ; `size` en centièmes de point. */
const textEntry = (size: number, lum = '65000', off = '35000', extra = '') =>
  `<cs:fontRef idx="minor"><a:schemeClr val="tx1"><a:lumMod val="${lum}"/><a:lumOff val="${off}"/></a:schemeClr></cs:fontRef>${extra}<cs:defRPr sz="${size}"/>`;
const refs = (fill = '<cs:fillRef idx="0"/>', ln = '<cs:lnRef idx="0"/>') => `${ln}${fill}<cs:effectRef idx="0"/>`;
const lineSpPr = (w: number, lum: string, off: string) => `<cs:spPr><a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="${lum}"/><a:lumOff val="${off}"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr>`;
const plain = (name: string) => `<cs:${name}>${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef></cs:${name}>`;

/** Style du graphique (partie cs:chartStyle), proche du style par défaut d'Excel 2016. */
export function chartStyleXml(): string {
  const auto = '<cs:fillRef idx="0"><cs:styleClr val="auto"/></cs:fillRef>';
  const autoLn = '<cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef>';
  const entries = [
    `<cs:axisTitle>${refs()}${textEntry(1000)}</cs:axisTitle>`,
    `<cs:categoryAxis>${refs()}${textEntry(900, '65000', '35000', lineSpPr(9525, '25000', '75000'))}</cs:categoryAxis>`,
    `<cs:chartArea>${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="25000"/><a:lumOff val="75000"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz="1330"/></cs:chartArea>`,
    `<cs:dataLabel>${refs()}${textEntry(900, '75000', '25000')}</cs:dataLabel>`,
    `<cs:dataLabelCallout>${refs()}<cs:fontRef idx="minor"><a:schemeClr val="dk1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln><a:solidFill><a:schemeClr val="dk1"><a:lumMod val="25000"/><a:lumOff val="75000"/></a:schemeClr></a:solidFill></a:ln></cs:spPr><cs:defRPr sz="900"/></cs:dataLabelCallout>`,
    `<cs:dataPoint>${refs(auto)}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr></cs:dataPoint>`,
    `<cs:dataPoint3D>${refs(auto)}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></cs:spPr></cs:dataPoint3D>`,
    `<cs:dataPointLine>${autoLn}<cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointLine>`,
    `<cs:dataPointMarker>${autoLn}${auto}<cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></cs:spPr></cs:dataPointMarker>`,
    '<cs:dataPointMarkerLayout symbol="circle" size="5"/>',
    `<cs:dataPointWireframe>${autoLn}<cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:spPr><a:ln w="9525" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointWireframe>`,
    `<cs:dataTable>${refs()}${textEntry(900, '65000', '35000', lineSpPr(9525, '15000', '85000'))}</cs:dataTable>`,
    plain('downBar'),
    plain('dropLine'),
    plain('errorBar'),
    plain('floor'),
    `<cs:gridlineMajor>${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>${lineSpPr(9525, '15000', '85000')}</cs:gridlineMajor>`,
    `<cs:gridlineMinor>${refs()}<cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef>${lineSpPr(9525, '5000', '95000')}</cs:gridlineMinor>`,
    plain('hiLoLine'),
    plain('leaderLine'),
    `<cs:legend>${refs()}${textEntry(900)}</cs:legend>`,
    plain('plotArea'),
    plain('plotArea3D'),
    `<cs:seriesAxis>${refs()}${textEntry(900)}</cs:seriesAxis>`,
    plain('seriesLine'),
    `<cs:title>${refs()}${textEntry(1400)}</cs:title>`,
    plain('trendline'),
    `<cs:trendlineLabel>${refs()}${textEntry(900)}</cs:trendlineLabel>`,
    plain('upBar'),
    `<cs:valueAxis>${refs()}${textEntry(900)}</cs:valueAxis>`,
    plain('wall'),
  ];
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cs:chartStyle xmlns:cs="${NS_CS}" xmlns:a="${NS_A}" id="395">${entries.join('')}</cs:chartStyle>`;
}

/** Palette du graphique (partie cs:colorStyle) : les couleurs d'accentuation du thème, comme Excel. */
export function chartColorsXml(): string {
  const accents = [1, 2, 3, 4, 5, 6].map((n) => `<a:schemeClr val="accent${n}"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cs:colorStyle xmlns:cs="${NS_CS}" xmlns:a="${NS_A}" meth="cycle" id="10">${accents}<cs:variation/><cs:variation><a:lumMod val="60000"/></cs:variation><cs:variation><a:lumMod val="80000"/><a:lumOff val="20000"/></cs:variation><cs:variation><a:lumMod val="80000"/></cs:variation><cs:variation><a:lumMod val="60000"/><a:lumOff val="40000"/></cs:variation><cs:variation><a:lumMod val="50000"/></cs:variation><cs:variation><a:lumMod val="70000"/><a:lumOff val="30000"/></cs:variation><cs:variation><a:lumMod val="70000"/></cs:variation><cs:variation><a:lumMod val="50000"/><a:lumOff val="50000"/></cs:variation></cs:colorStyle>`;
}
