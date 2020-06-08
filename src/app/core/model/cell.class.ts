import { CellStyle } from './cell-style.class';

export class Cell {
  // cells: Cell[][];
  // offsetLeft: number;
  // offsetTop: number;
  columns?: { x?: number; width?: number }[];
  rows?: { y?: number; height?: number }[];
  position: { row: number; column: number };
  x: number;
  y: number;
  width: number;
  height: number;
  content: any;
  style: CellStyle;
  type: string;
  rowSpan: number;
  colSpan: number;
  isCombined?: boolean;
  combineCell?: Cell;
  hidden?: boolean;
}
