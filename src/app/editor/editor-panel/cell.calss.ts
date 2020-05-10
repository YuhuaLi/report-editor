export default class Cell {
  x: number;
  y: number;
  width: number;
  height: number;
  content: any;
  background: string;
  color: string;
  fontWeight: string;
  textAlign: 'left' | 'right' | 'center' | 'start' | 'end';
  textBaseline:
    | 'top'
    | 'hanging'
    | 'middle'
    | 'alphabetic'
    | 'ideographic'
    | 'bottom';
  position: { row: number; column: number };
}
