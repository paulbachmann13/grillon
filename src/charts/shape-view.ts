/**
 * Formes et zones de texte dans la feuille : « DOM flottant » Univer (déplaçable, redimensionnable,
 * enregistré avec le classeur et exporté dans le .xlsx). Double-clic ou ⚙ pour modifier.
 */
import { IconManager } from '@univerjs/preset-sheets-core';
import { CommandType, ICommandService, type FUniver } from '@univerjs/presets';
import type { FWorksheet } from '@univerjs/sheets/facade';
import { createElement, type ComponentType, type ReactNode } from 'react';
import { t } from '../i18n';
import { box, checkbox, field, openDialog, select } from '../ui/dialog';
import { workbook } from './chart-store';
import { resolveShape, SHAPE_DEFAULT_SIZE, SHAPE_FILL, SHAPE_LABELS, SHAPE_STROKE, type ShapeKind, type ShapeSpec } from './shape';

export const SHAPE_COMPONENT = 'GrillonShape';

export interface ShapeData {
  shapeId: string;
  spec: ShapeSpec;
}

interface FloatDom {
  id: string;
  componentKey: string;
  data?: ShapeData;
}

const floatDoms = (ws: FWorksheet) => (ws as unknown as { getAllFloatDoms(): FloatDom[] }).getAllFloatDoms();

function findShape(api: FUniver, shapeId: string): { ws: FWorksheet; dom: FloatDom } | null {
  const wb = workbook(api);
  if (!wb) return null;
  for (const ws of wb.getSheets() as unknown as FWorksheet[]) {
    const dom = floatDoms(ws).find((d) => d.componentKey === SHAPE_COMPONENT && d.data?.shapeId === shapeId);
    if (dom) return { ws, dom };
  }
  return null;
}

const markerId = (id: string) => `g-arrow-${id}`;

export function createShapeComponent(api: FUniver): ComponentType<{ data?: ShapeData }> {
  return function ShapeView(props: { data?: ShapeData }) {
    const spec = props.data?.spec;
    if (!spec || !props.data) return null;
    const r = resolveShape(spec);
    const id = props.data.shapeId;
    const stroke = r.stroke === 'none' ? 'none' : r.stroke;
    const common = { fill: r.fill === 'none' ? 'none' : r.fill, stroke, strokeWidth: r.strokeWidth, vectorEffect: 'non-scaling-stroke' };
    let geometry: ReactNode;
    if (spec.kind === 'ellipse') geometry = createElement('ellipse', { cx: 50, cy: 50, rx: 49.5, ry: 49.5, ...common });
    else if (spec.kind === 'line' || spec.kind === 'arrow') geometry = createElement('line', { x1: 2, y1: 2, x2: 98, y2: 98, ...common, fill: 'none', markerEnd: spec.kind === 'arrow' ? `url(#${markerId(id)})` : undefined });
    else geometry = createElement('rect', { x: 0.5, y: 0.5, width: 99, height: 99, rx: spec.kind === 'roundRect' ? 12 : 0, ...common });
    const svg = createElement(
      'svg',
      { viewBox: '0 0 100 100', preserveAspectRatio: 'none', className: 'g-shape-svg', 'aria-hidden': true },
      spec.kind === 'arrow'
        ? createElement('defs', null, createElement('marker', { id: markerId(id), viewBox: '0 0 10 10', refX: 8, refY: 5, markerWidth: 7, markerHeight: 7, orient: 'auto-start-reverse' }, createElement('path', { d: 'M0 0L10 5L0 10z', fill: stroke })))
        : null,
      geometry,
    );
    const label =
      spec.kind !== 'line' && spec.kind !== 'arrow' && spec.text
        ? createElement(
            'div',
            {
              className: 'g-shape-text',
              style: { color: r.color, fontSize: `${r.fontSize}pt`, fontWeight: spec.bold ? 700 : 400, fontStyle: spec.italic ? 'italic' : 'normal', textAlign: r.align, justifyContent: r.align === 'center' ? 'center' : r.align === 'right' ? 'flex-end' : 'flex-start' },
            },
            spec.text,
          )
        : null;
    return createElement(
      'div',
      { className: 'g-shape', 'data-testid': 'shape', onDoubleClick: () => void shapeDialog(api, id) },
      svg,
      label,
      createElement(
        'button',
        { type: 'button', className: 'g-chart-edit', title: t('shapes.edit'), 'aria-label': t('shapes.edit'), onPointerDown: (e: Event) => e.stopPropagation(), onClick: () => void shapeDialog(api, id) },
        '⚙',
      ),
    );
  };
}

function updateShape(api: FUniver, shapeId: string, spec: ShapeSpec): void {
  const found = findShape(api, shapeId);
  if (found) (found.ws as unknown as { updateFloatDom(id: string, c: object): void }).updateFloatDom(found.dom.id, { data: { shapeId, spec } });
}

function removeShape(api: FUniver, shapeId: string): void {
  const found = findShape(api, shapeId);
  if (found) (found.ws as unknown as { removeFloatDom(id: string): void }).removeFloatDom(found.dom.id);
}

/** Insère une forme près de la cellule active. */
export function insertShape(api: FUniver, kind: ShapeKind): string | null {
  const wb = workbook(api);
  const ws = wb?.getActiveSheet() as unknown as FWorksheet | undefined;
  const range = wb?.getActiveRange();
  if (!ws || !range) return null;
  const config = (ws as unknown as { getSheet(): { getConfig(): { rowHeader?: { width?: number }; columnHeader?: { height?: number } } } }).getSheet().getConfig();
  let left = config.rowHeader?.width ?? 46;
  for (let c = 0; c < range.getColumn(); c++) left += ws.getColumnWidth(c);
  let top = config.columnHeader?.height ?? 20;
  for (let r = 0; r < range.getRow(); r++) top += ws.getRowHeight(r);
  const size = SHAPE_DEFAULT_SIZE[kind];
  const shapeId = `shape-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
  const spec: ShapeSpec = { kind, ...(kind === 'textbox' ? { text: t('shapes.textboxDefault') } : {}) };
  (ws as unknown as { addFloatDomToPosition(l: object): unknown }).addFloatDomToPosition({
    componentKey: SHAPE_COMPONENT,
    initPosition: { startX: left + 12, endX: left + 12 + size.w, startY: top + 12, endY: top + 12 + size.h },
    data: { shapeId, spec } satisfies ShapeData,
    allowTransform: true,
  });
  return shapeId;
}

/** Boîte « Format de la forme ». */
export async function shapeDialog(api: FUniver, shapeId: string): Promise<void> {
  const found = findShape(api, shapeId);
  if (!found?.dom.data) return;
  const spec: ShapeSpec = structuredClone(found.dom.data.spec);
  const r = resolveShape(spec);
  const linear = spec.kind === 'line' || spec.kind === 'arrow';
  const text = document.createElement('textarea');
  text.rows = 3;
  text.value = spec.text ?? '';
  const colorInput = (value: string) => Object.assign(document.createElement('input'), { type: 'color', value });
  const noFill = checkbox(t('shapes.noFill'), r.fill === 'none');
  const fill = colorInput(r.fill === 'none' ? SHAPE_FILL : r.fill);
  const noStroke = checkbox(t('shapes.noStroke'), r.stroke === 'none');
  const stroke = colorInput(r.stroke === 'none' ? SHAPE_STROKE : r.stroke);
  const width = Object.assign(document.createElement('input'), { type: 'number', step: 'any', min: '0', value: String(r.strokeWidth) });
  const size = Object.assign(document.createElement('input'), { type: 'number', step: '1', min: '6', max: '96', value: String(r.fontSize) });
  const textColor = colorInput(r.color);
  const bold = checkbox(t('shapes.bold'), !!spec.bold);
  const italic = checkbox(t('shapes.italic'), !!spec.italic);
  const align = select([['left', t('shapes.alignLeft')], ['center', t('shapes.alignCenter')], ['right', t('shapes.alignRight')]], r.align);
  let removed = false;
  const del = document.createElement('button');
  del.type = 'button';
  del.className = 'g-danger';
  del.textContent = t('shapes.delete');
  del.addEventListener('click', () => {
    removed = true;
    removeShape(api, shapeId);
    (del.closest('.g-dialog')?.querySelector('.g-dialog-actions button:not(.primary)') as HTMLButtonElement | null)?.click();
  });
  const rows: HTMLElement[] = [];
  if (!linear) rows.push(field(t('shapes.text'), text));
  if (!linear) rows.push(noFill.el, field(t('shapes.fill'), fill));
  rows.push(noStroke.el, field(t('shapes.stroke'), stroke), field(t('shapes.width'), width));
  if (!linear) rows.push(field(t('shapes.fontSize'), size), field(t('shapes.textColor'), textColor), bold.el, italic.el, field(t('shapes.align'), align));
  rows.push(del);
  const ok = await openDialog({ title: t('shapes.title', { name: SHAPE_LABELS[spec.kind] }), testId: 'dialog-shape', body: box(...rows), onOk: () => true });
  if (!ok || removed) return;
  const next: ShapeSpec = { kind: spec.kind };
  if (text.value && !linear) next.text = text.value;
  const defaults = resolveShape({ kind: spec.kind });
  if (!linear) {
    const wanted = noFill.input.checked ? 'none' : fill.value;
    if (wanted !== defaults.fill) next.fill = wanted;
  }
  const strokeWanted = noStroke.input.checked ? 'none' : stroke.value;
  if (strokeWanted !== defaults.stroke) next.stroke = strokeWanted;
  const w = Number(width.value);
  if (Number.isFinite(w) && w !== defaults.strokeWidth) next.strokeWidth = w;
  if (!linear) {
    const fs = Number(size.value);
    if (Number.isFinite(fs) && fs !== defaults.fontSize) next.fontSize = fs;
    if (textColor.value.toUpperCase() !== defaults.color.toUpperCase()) next.color = textColor.value;
    if (bold.input.checked) next.bold = true;
    if (italic.input.checked) next.italic = true;
    if (align.value !== defaults.align) next.align = align.value as ShapeSpec['align'];
  }
  updateShape(api, shapeId, next);
}

function icon(path: string): ComponentType<Record<string, unknown>> {
  return (props) =>
    createElement('svg', { viewBox: '0 0 16 16', width: '1em', height: '1em', fill: 'none', stroke: 'currentColor', strokeWidth: 1.3, strokeLinecap: 'round', strokeLinejoin: 'round', ...props }, createElement('path', { d: path }));
}

export function installShapes(api: FUniver): void {
  const injector = (api as unknown as { _injector: { get<T>(id: unknown): T } })._injector;
  const commands = injector.get<ICommandService>(ICommandService);
  injector.get<IconManager>(IconManager).register({ GrillonShapesIcon: icon('M2 9h6v5H2zM11 3a3 3 0 110 6 3 3 0 010-6zM9 14l5-4') } as never);
  const id = 'grillon.insert.shape';
  const path = ['ribbon.insert', 'ribbon.insert.media'];
  api.createSubmenu({ id, icon: 'GrillonShapesIcon', title: t('shapes.menu'), tooltip: t('shapes.menu') }).appendTo(path);
  (Object.keys(SHAPE_LABELS) as ShapeKind[]).forEach((kind, order) => {
    const itemId = `${id}.${kind}`;
    if (!commands.hasCommand(itemId)) commands.registerCommand({ id: itemId, type: CommandType.COMMAND, handler: () => insertShape(api, kind) !== null });
    api.createMenu({ id: itemId, title: SHAPE_LABELS[kind], action: itemId, order }).appendTo([...path, id]);
  });
}
