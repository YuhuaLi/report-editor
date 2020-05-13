export default class Cell {
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
