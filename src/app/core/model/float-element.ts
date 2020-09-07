export class FloatElement {
  constructor(x, y, width, height, content) {
    this.x = x;
    this.y = y;
    this.width = width;
    this.height = height;
    this.content = content;
    this.isActive = true;
  }

  x: number;
  y: number;
  width: number;
  height: number;
  content: any;
  isActive: boolean;
}
