/**
 * Formes et zones de texte (modèle) : rectangle, rectangle arrondi, ellipse, ligne, flèche, zone de texte.
 */
export type ShapeKind = 'rect' | 'roundRect' | 'ellipse' | 'line' | 'arrow' | 'textbox';

export interface ShapeSpec {
  kind: ShapeKind;
  text?: string;
  /** Couleur de remplissage (« none » : transparent). */
  fill?: string;
  /** Couleur du contour (« none » : sans contour). */
  stroke?: string;
  /** Épaisseur du contour en points. */
  strokeWidth?: number;
  fontSize?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  align?: 'left' | 'center' | 'right';
}

export const SHAPE_FILL = '#4472C4';
export const SHAPE_STROKE = '#2F528F';
export const TEXTBOX_STROKE = '#7F7F7F';

export interface ResolvedShape {
  fill: string;
  stroke: string;
  strokeWidth: number;
  fontSize: number;
  color: string;
  align: 'left' | 'center' | 'right';
}

/** Valeurs par défaut selon le type (comme Excel : forme bleue à texte blanc, zone de texte transparente). */
export function resolveShape(s: ShapeSpec): ResolvedShape {
  const text = s.kind === 'textbox';
  const linear = s.kind === 'line' || s.kind === 'arrow';
  return {
    fill: s.fill ?? (text || linear ? 'none' : SHAPE_FILL),
    stroke: s.stroke ?? (text ? TEXTBOX_STROKE : SHAPE_STROKE),
    strokeWidth: s.strokeWidth ?? (linear ? 1.5 : 1),
    fontSize: s.fontSize ?? 11,
    color: s.color ?? (text ? '#000000' : '#FFFFFF'),
    align: s.align ?? (text ? 'left' : 'center'),
  };
}

export const SHAPE_LABELS: Record<ShapeKind, string> = {
  rect: 'Rectangle',
  roundRect: 'Rectangle aux coins arrondis',
  ellipse: 'Ellipse',
  line: 'Ligne',
  arrow: 'Flèche',
  textbox: 'Zone de texte',
};

/** Taille initiale (px) à l'insertion. */
export const SHAPE_DEFAULT_SIZE: Record<ShapeKind, { w: number; h: number }> = {
  rect: { w: 160, h: 80 },
  roundRect: { w: 160, h: 80 },
  ellipse: { w: 140, h: 90 },
  line: { w: 140, h: 60 },
  arrow: { w: 140, h: 60 },
  textbox: { w: 180, h: 60 },
};
