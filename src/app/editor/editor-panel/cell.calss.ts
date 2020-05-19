export default class Cell {
  // cells: Cell[][];
  // offsetLeft: number;
  // offsetTop: number;
  columns?: { x?: number; width?: number }[];
  rows?: { y?: number; height?: number }[];
  x: number;
  y: number;
  width: number;
  height: number;
  content: any;
  background: string;
  color: string;
  fontWeight: string;
  fontStyle: string;
  fontSize: number;
  fontFamily: string;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  position: { row: number; column: number };
  type: string;
  borderWidth: number;
  borderColor: string;
}
