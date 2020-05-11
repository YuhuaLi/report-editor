import { Style } from './style.enum';
import {
  Component,
  OnInit,
  ViewChild,
  ElementRef,
  AfterViewInit,
} from '@angular/core';
import Cell from './cell.calss';

@Component({
  selector: 'app-editor-panel',
  templateUrl: './editor-panel.component.html',
  styleUrls: ['./editor-panel.component.scss'],
})
export class EditorPanelComponent implements OnInit, AfterViewInit {
  @ViewChild('panel') panel: ElementRef;

  width = 0;
  height = 0;
  workCells = [];
  ViewOrigin = { x: 0, y: 0 };
  viewRowCount = 0;
  viewColumnCount = 0;
  startRowIndex = 0;
  startColumnIndex = 0;
  offsetWidth = 0;
  offsetHeight = 30;
  cells: Cell[][] = [];
  viewCells: Cell[][] = [];
  activeCell: Cell;
  activeRange: {
    rowStart: number;
    rowEnd?: number;
    columnStart: number;
    columnEnd?: number;
  };
  state: any = {};
  isTicking = false;

  ctx: CanvasRenderingContext2D;

  constructor(private elmentRef: ElementRef) {}

  ngOnInit(): void {
    this.width = this.elmentRef.nativeElement.offsetWidth;
    this.height = this.elmentRef.nativeElement.offsetHeight;
  }

  ngAfterViewInit() {
    this.ctx = this.panel.nativeElement.getContext('2d');
    this.viewRowCount =
      Math.ceil((this.height - this.offsetHeight) / Style.cellHeight) + 2;
    this.ctx.font = `${Style.rulerCellFontWeight} ${Style.rulerCellFontSize}px ${Style.rulerCellFontFamily}`;
    this.offsetWidth = Math.ceil(
      this.ctx.measureText(`  ${this.viewRowCount}  `).width
    );
    this.viewColumnCount =
      Math.ceil((this.width - this.offsetWidth) / Style.cellWidth) + 2;

    this.cells = Array.from({ length: this.viewRowCount + 1 }, (rv, rk) => {
      return Array.from({ length: this.viewColumnCount + 1 }, (cv, ck) => {
        const isXRuler = rk === 0;
        const isYRuler = ck === 0;
        return {
          position: { row: rk, column: ck },
          x: ck === 0 ? 0 : this.offsetWidth + (ck - 1) * Style.cellWidth,
          y: isXRuler ? 0 : this.offsetHeight + (rk - 1) * Style.cellHeight,
          width: isYRuler ? this.offsetWidth : Style.cellWidth,
          height: isXRuler ? this.offsetHeight : Style.cellHeight,
          type: isXRuler || isYRuler ? 'ruler' : 'cell',
          content: {
            value:
              isYRuler && !isXRuler
                ? this.generateRowNum(rk)
                : isXRuler && !isYRuler
                ? this.generateColumnNum(ck)
                : null,
          },
          fontWeight:
            isXRuler || isYRuler
              ? Style.rulerCellFontWeight
              : Style.cellFontWeight,
          textAlign: Style.cellTextAlign,
          textBaseline: Style.cellTextBaseline,
          fontStyle: Style.cellFontStyle,
          fontFamily:
            isXRuler || isYRuler
              ? Style.rulerCellFontFamily
              : Style.cellFontFamily,
          fontSize:
            isXRuler || isYRuler ? Style.rulerCellFontSize : Style.cellFontSize,
          background:
            (isXRuler && !isYRuler) || (isYRuler && !isXRuler)
              ? Style.rulerCellBackgroundColor
              : Style.cellBackgroundColor,
          color: isXRuler || isYRuler ? Style.rulerCellColor : Style.cellColor,
          borderWidth:
            isXRuler || isYRuler
              ? Style.rulerCellBorderWidth
              : Style.cellBorderWidth,
          borderColor:
            isXRuler || isYRuler
              ? Style.rulerCellBorderColor
              : Style.cellBorderColor,
        };
      }).reduce((acc, cur) => acc.push(cur) && acc, []);
    });
    console.log(this.cells);
    // this.viewCells = this.cells.slice(1).map((cells) => cells.slice(1));
    this.viewCells = this.cells;
    this.drawPanel();
    this.drawRuler();
    this.drawSrcollBar();
  }

  drawSrcollBar() {
    this.ctx.save();
    this.ctx.fillStyle = Style.scrollBarBackgroundColor;
    this.ctx.strokeStyle = Style.scrollBarBorderColor;
    this.ctx.lineWidth = Style.scrollBarBorderWidth;
    this.ctx.fillRect(
      0,
      this.height - Style.scrollBarWidth,
      this.width - Style.scrollBarWidth,
      this.height
    );
    this.ctx.strokeRect(
      0,
      this.height - Style.scrollBarWidth,
      this.width - Style.scrollBarWidth,
      this.height
    );
    this.ctx.fillRect(
      this.width - Style.scrollBarWidth,
      0,
      this.width,
      this.height - Style.scrollBarWidth
    );
    this.ctx.strokeRect(
      this.width - Style.scrollBarWidth,
      0,
      this.width,
      this.height - Style.scrollBarWidth
    );
    this.ctx.fillRect(
      this.width - Style.scrollBarWidth,
      this.height - Style.scrollBarWidth,
      this.width,
      this.height
    );

    this.ctx.fillStyle = Style.scrollBarThumbColor;
    this.roundedRect(
      this.width - Style.scrollBarWidth + Style.scrollBarThumbMargin,
      Style.scrollBarBorderWidth,
      Style.scrollBarWidth - 2 * Style.scrollBarThumbMargin,
      100,
      Style.scrollBarThumbRadius
    );

    this.ctx.restore();
  }

  roundedRect(x, y, width, height, radius) {
    this.ctx.beginPath();
    this.ctx.moveTo(x, y + radius);
    this.ctx.lineTo(x, y + height - radius);
    this.ctx.quadraticCurveTo(x, y + height, x + radius, y + height);
    this.ctx.lineTo(x + width - radius, y + height);
    this.ctx.quadraticCurveTo(
      x + width,
      y + height,
      x + width,
      y + height - radius
    );
    this.ctx.lineTo(x + width, y + radius);
    this.ctx.quadraticCurveTo(x + width, y, x + width - radius, y);
    this.ctx.lineTo(x + radius, y);
    this.ctx.quadraticCurveTo(x, y, x, y + radius);
    this.ctx.fill();
  }

  drawRuler() {
    this.ctx.save();
    const columns = this.viewCells[0].slice(1);
    const rows = this.viewCells.slice(1).map((cells) => cells[0]);
    this.ctx.fillStyle = columns[0].background;
    this.ctx.textAlign = columns[0].textAlign;
    this.ctx.textBaseline = columns[0].textBaseline;
    if (
      this.ctx.font !==
      `${columns[0].fontStyle} ${columns[0].fontWeight} ${columns[0].fontSize}px ${columns[0].fontFamily}`
    ) {
      this.ctx.font = `${columns[0].fontStyle} ${columns[0].fontWeight} ${columns[0].fontSize}px ${columns[0].fontFamily}`;
      console.log(
        this.ctx.font,
        `${columns[0].fontStyle} ${columns[0].fontWeight} ${columns[0].fontSize}px ${columns[0].fontFamily}`
      );
    }
    for (let len = columns.length, i = len - 1; i >= 0; i--) {
      this.ctx.fillRect(
        columns[i].x,
        columns[i].y,
        columns[i].width,
        columns[i].height
      );
    }
    for (let len = rows.length, i = len - 1; i >= 0; i--) {
      this.ctx.fillRect(rows[i].x, rows[i].y, rows[i].width, rows[i].height);
    }
    this.ctx.fillStyle = columns[0].color;
    this.ctx.strokeStyle = Style.rulerCellBorderColor;
    for (let len = columns.length, i = len - 1; i >= 0; i--) {
      this.ctx.strokeRect(
        columns[i].x,
        columns[i].y,
        columns[i].width,
        columns[i].height
      );
      if (columns[i].content.value) {
        this.ctx.fillText(
          columns[i].content.value,
          columns[i].x + columns[i].width / 2,
          columns[i].y + columns[i].height / 2,
          columns[i].width - 2 * columns[0].borderWidth
        );
      }
    }
    for (let len = rows.length, i = len - 1; i >= 0; i--) {
      this.ctx.strokeRect(rows[i].x, rows[i].y, rows[i].width, rows[i].height);
      if (rows[i].content.value) {
        this.ctx.fillText(
          rows[i].content.value,
          rows[i].x + rows[i].width / 2,
          rows[i].y + rows[i].height / 2,
          rows[i].width - 2 * columns[0].borderWidth
        );
      }
    }
    this.ctx.restore();
  }

  drawPanel() {
    this.ctx.save();
    this.ctx.clearRect(0, 0, this.width, this.height);
    for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
      for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
        this.drawCell(this.viewCells[i][j]);
      }
    }
    this.ctx.restore();
  }

  drawCell(cell: Cell) {
    this.ctx.clearRect(cell.x, cell.y, cell.width, cell.height);
    if (cell.background) {
      if (this.ctx.fillStyle !== cell.background) {
        this.ctx.fillStyle = cell.background;
      }
      this.ctx.fillRect(cell.x, cell.y, cell.width, cell.height);
    }
    if (this.ctx.strokeStyle !== cell.borderColor) {
      this.ctx.strokeStyle = cell.borderColor;
    }
    this.ctx.strokeRect(cell.x, cell.y, cell.width, cell.height);
    if (cell.content.value) {
      if (this.ctx.fillStyle !== cell.color) {
        this.ctx.fillStyle = cell.color;
      }
      if (cell.textAlign && this.ctx.textAlign !== cell.textAlign) {
        this.ctx.textAlign = cell.textAlign;
      }
      if (cell.textBaseline && this.ctx.textBaseline !== cell.textBaseline) {
        this.ctx.textBaseline = cell.textBaseline;
      }
      if (
        this.ctx.font !==
        `${cell.fontStyle} ${cell.fontWeight} ${cell.fontSize}px ${cell.fontFamily}`
      ) {
        this.ctx.font = `${cell.fontStyle} ${cell.fontWeight} ${cell.fontSize}px ${cell.fontFamily}`;
      }
      this.ctx.fillText(
        cell.content.value,
        cell.x + cell.width / 2,
        cell.y + cell.height / 2,
        cell.width - 2 * cell.borderWidth
      );
    }
  }

  setActive(activeRange?: {
    rowStart: number;
    rowEnd: number;
    columnStart: number;
    columnEnd: number;
  }) {
    this.ctx.save();
    if (this.activeRange) {
      const rowStart = Math.min(
        this.activeRange.rowEnd,
        this.activeRange.rowStart
      );
      const rowEnd = Math.max(
        this.activeRange.rowEnd,
        this.activeRange.rowStart
      );
      const columnStart = Math.min(
        this.activeRange.columnStart,
        this.activeRange.columnEnd
      );
      const columnEnd = Math.max(
        this.activeRange.columnStart,
        this.activeRange.columnEnd
      );
      for (let i = rowStart - 1; i < rowEnd + 2; i++) {
        for (let j = columnStart - 1; j < columnEnd + 2; j++) {
          if (
            i > rowStart + 1 &&
            i < rowEnd - 1 &&
            j > columnStart + 1 &&
            j < columnEnd - 1
          ) {
            continue;
          }
          this.drawCell(this.viewCells[i][j]);
        }
      }
    }
    this.activeRange = activeRange;
    if (activeRange) {
      const rowStart = Math.min(
        this.activeRange.rowEnd,
        this.activeRange.rowStart
      );
      const rowEnd = Math.max(
        this.activeRange.rowEnd,
        this.activeRange.rowStart
      );
      const columnStart = Math.min(
        this.activeRange.columnStart,
        this.activeRange.columnEnd
      );
      const columnEnd = Math.max(
        this.activeRange.columnStart,
        this.activeRange.columnEnd
      );
      this.ctx.strokeStyle = Style.activeCellBorderColor;
      this.ctx.shadowColor = Style.activeCellShadowColor;
      this.ctx.shadowBlur = Style.activeCellShadowBlur;
      this.ctx.strokeRect(
        this.viewCells[rowStart][columnStart].x,
        this.viewCells[rowStart][columnStart].y,
        this.viewCells[rowEnd][columnEnd].x -
          this.viewCells[rowStart][columnStart].x +
          this.viewCells[rowEnd][columnEnd].width,
        this.viewCells[rowEnd][columnEnd].y -
          this.viewCells[rowStart][columnStart].y +
          this.viewCells[rowEnd][columnEnd].height
      );
    }
    this.ctx.restore();
  }

  generateColumnNum(columnIndex: number): string {
    const columnNumArr: string[] = [];
    while (columnIndex > 26) {
      columnNumArr.unshift(String.fromCharCode((columnIndex % 26 || 26) + 64));
      columnIndex = Math.floor(columnIndex / 26);
    }
    columnNumArr.unshift(String.fromCharCode(columnIndex + 64));

    return columnNumArr.join('');
  }

  generateRowNum(rowIndex: number): string {
    return rowIndex > 0 ? rowIndex.toString() : '';
  }

  onMouseOver(event: MouseEvent) {
    this.state.isMouseDown = false;
    // console.log('mouseover', event);
  }

  onMouseEnter(event) {
    // console.log('mouseenter', event);
  }

  onContextMenu(event: MouseEvent) {
    event.returnValue = false;
  }

  onMouseDown(event: MouseEvent) {
    if (event.button === 2) {
      event.returnValue = false;
      return;
    }
    this.state.isMouseDown = true;
    if (
      (event.clientX < this.offsetWidth && event.clientY > this.offsetHeight) ||
      (event.clientX > this.offsetWidth && event.clientY < this.offsetHeight)
    ) {
      return;
    }
    const canvas = document.createElement('canvas');
    canvas.width = this.width;
    canvas.height = this.height;
    const ctx = canvas.getContext('2d');
    for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
      for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
        const cell = this.viewCells[i][j];
        ctx.clearRect(0, 0, this.width, this.height);
        ctx.rect(cell.x, cell.y, cell.width, cell.height);
        if (ctx.isPointInPath(event.clientX, event.clientY)) {
          this.setActive({
            rowStart: cell.position.row,
            columnStart: cell.position.column,
            rowEnd: cell.position.row,
            columnEnd: cell.position.column,
          });
          return;
        }
      }
    }
  }

  // @throttle(20)
  onMouseMove(event: MouseEvent) {
    if (event.clientX < this.offsetWidth || event.clientY < this.offsetHeight) {
      this.panel.nativeElement.style.cursor = 'default';
    } else {
      this.panel.nativeElement.style.cursor = 'cell';
    }
    if (this.state.isMouseDown) {
      if (!this.isTicking) {
        window.requestAnimationFrame(() => {
          this.calcActive(event.clientX, event.clientY);
          this.isTicking = false;
        });
      }
      this.isTicking = true;
    }
  }

  calcActive(x: number, y: number) {
    const canvas = document.createElement('canvas');
    canvas.width = this.width;
    canvas.height = this.height;
    const ctx = canvas.getContext('2d');
    for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
      for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
        const cell = this.viewCells[i][j];
        ctx.clearRect(0, 0, this.width, this.height);
        ctx.rect(cell.x, cell.y, cell.width, cell.height);
        if (ctx.isPointInPath(x, y)) {
          this.setActive({
            rowStart: this.activeRange.rowStart,
            columnStart: this.activeRange.columnStart,
            rowEnd: cell.position.row,
            columnEnd: cell.position.column,
          });
          return;
        }
      }
    }
  }

  onMouseUp(event: MouseEvent) {
    this.state.isMouseDown = false;
  }

  onWheel(event) {
    console.log('wheel', event);
  }
}
