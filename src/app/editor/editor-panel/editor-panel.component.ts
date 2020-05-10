import { Constant } from './constant.enum';
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
  columns = [];
  rows = [];
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

    this.columns = Array.from(
      {
        length: Math.ceil(this.width / Constant.INIT_COLUMN_WIDTH),
      },
      (v, k) => ({
        x: Constant.INIT_COLUMN_WIDTH * k,
        y: 0,
        width: Constant.INIT_COLUMN_WIDTH,
        height: Constant.INIT_ROW_HEIGHT,
        content: { text: this.generateColumnNum(k) },
      })
    );
  }

  ngAfterViewInit() {
    this.ctx = this.panel.nativeElement.getContext('2d');
    this.ctx.strokeStyle = Constant.LINE_COLOR;
    this.ctx.lineWidth = Constant.LINE_WIDTH;
    this.ctx.font = `${Constant.FONT_SIZE}px ${Constant.FONT_STYLE}`;
    this.viewRowCount =
      Math.ceil((this.height - this.offsetHeight) / Constant.INIT_ROW_HEIGHT) +
      2;
    this.offsetWidth = Math.ceil(
      this.ctx.measureText(`  ${this.viewRowCount}  `).width
    );
    this.viewColumnCount =
      Math.ceil((this.width - this.offsetWidth) / Constant.INIT_COLUMN_WIDTH) +
      2;

    this.cells = Array.from({ length: this.viewRowCount + 1 }, (rv, rk) => {
      return Array.from({ length: this.viewColumnCount + 1 }, (cv, ck) => ({
        position: { row: rk, column: ck },
        x:
          ck === 0
            ? 0
            : this.offsetWidth + (ck - 1) * Constant.INIT_COLUMN_WIDTH,
        y:
          rk === 0
            ? 0
            : this.offsetHeight + (rk - 1) * Constant.INIT_ROW_HEIGHT,
        width: ck === 0 ? this.offsetWidth : Constant.INIT_COLUMN_WIDTH,
        height: rk === 0 ? this.offsetHeight : Constant.INIT_ROW_HEIGHT,
        content: {
          value:
            ck === 0 && rk !== 0
              ? this.generateRowNum(rk)
              : ck !== 0 && rk === 0
              ? this.generateColumnNum(ck)
              : null,
        },
        fontWeight: 'bold',
        textAlign: 'center',
        textBaseline: 'middle',
        fontStyle: 'normal',
        background:
          (ck === 0 && rk !== 0) || (ck !== 0 && rk === 0)
            ? Constant.BACK_COLOR
            : null,
        color: Constant.BLACK_COLOR,
      })).reduce((acc, cur) => acc.push(cur) && acc, []);
    });
    console.log(this.cells);
    this.viewCells = this.cells.slice(1).map((cells) => cells.slice(1));
    this.drawPanel();
    this.drawRuler();
  }

  drawRuler() {
    this.ctx.save();
    const columns = this.cells[0].slice(1);
    const rows = this.cells.slice(1).map((cells) => cells[0]);
    this.ctx.fillStyle = columns[0].background;
    this.ctx.textAlign = columns[0].textAlign;
    this.ctx.textBaseline = columns[0].textBaseline;
    if (
      columns[0].fontWeight &&
      this.ctx.font !==
        `${columns[0].fontWeight} ${Constant.FONT_SIZE}px ${Constant.FONT_STYLE}`
    ) {
      this.ctx.font = `${columns[0].fontWeight} ${Constant.FONT_SIZE}px ${Constant.FONT_STYLE}`;
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
    this.ctx.fillStyle = columns[0].color || Constant.BLACK_COLOR;
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
          columns[i].width - 2 * Constant.LINE_WIDTH
        );
      }
    }
    this.ctx.fillStyle = Constant.BLACK_COLOR;
    for (let len = rows.length, i = len - 1; i >= 0; i--) {
      this.ctx.strokeRect(rows[i].x, rows[i].y, rows[i].width, rows[i].height);
      if (rows[i].content.value) {
        this.ctx.fillText(
          rows[i].content.value,
          rows[i].x + rows[i].width / 2,
          rows[i].y + rows[i].height / 2,
          rows[i].width - 2 * Constant.LINE_WIDTH
        );
      }
    }
    this.ctx.restore();
  }

  drawPanel() {
    this.ctx.save();
    this.ctx.clearRect(0, 0, this.width, this.height);
    for (let rLen = this.viewCells.length, i = rLen - 1; i > -1; i--) {
      for (let cLen = this.viewCells[i].length, j = cLen - 1; j > -1; j--) {
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
    this.ctx.strokeRect(cell.x, cell.y, cell.width, cell.height);
    if (cell.content.value) {
      if (this.ctx.fillStyle !== (cell.color || Constant.BLACK_COLOR)) {
        this.ctx.fillStyle = cell.color || Constant.BLACK_COLOR;
      }
      if (cell.textAlign && this.ctx.textAlign !== cell.textAlign) {
        this.ctx.textAlign = cell.textAlign;
      }
      if (cell.textBaseline && this.ctx.textBaseline !== cell.textBaseline) {
        this.ctx.textBaseline = cell.textBaseline;
      }
      if (
        cell.fontWeight &&
        this.ctx.font !==
          `${cell.fontWeight} ${Constant.FONT_SIZE}px ${Constant.FONT_STYLE}`
      ) {
        this.ctx.font = `${cell.fontWeight} ${Constant.FONT_SIZE}px ${Constant.FONT_STYLE}`;
      }
      this.ctx.fillText(
        cell.content.value,
        cell.x + cell.width / 2,
        cell.y + cell.height / 2,
        cell.width - 2 * Constant.LINE_WIDTH
      );
    }
  }

  setActive(activeRange?: {
    rowStart: number;
    rowEnd: number;
    columnStart: number;
    columnEnd: number;
  }) {
    console.log(this.activeRange);
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
          this.drawCell(this.cells[i][j]);
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
      this.ctx.strokeStyle = Constant.ACTIVE_LINE_COLOR;
      this.ctx.shadowColor = Constant.BLACK_COLOR;
      this.ctx.shadowBlur = 5;
      this.ctx.strokeRect(
        this.cells[rowStart][columnStart].x,
        this.cells[rowStart][columnStart].y,
        this.cells[rowEnd][columnEnd].x -
          this.cells[rowStart][columnStart].x +
          this.cells[rowEnd][columnEnd].width,
        this.cells[rowEnd][columnEnd].y -
          this.cells[rowStart][columnStart].y +
          this.cells[rowEnd][columnEnd].height
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

  onMouseDown(event: MouseEvent) {
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
    for (let rLen = this.viewCells.length, i = rLen - 1; i > -1; i--) {
      for (let cLen = this.viewCells[i].length, j = cLen - 1; j > -1; j--) {
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
    console.log('not found');
  }

  onMouseMove(event: MouseEvent) {
    if (
      (event.clientX < this.offsetWidth && event.clientY > this.offsetHeight) ||
      (event.clientX > this.offsetWidth && event.clientY < this.offsetHeight)
    ) {
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
    for (let rLen = this.viewCells.length, i = rLen - 1; i > -1; i--) {
      for (let cLen = this.viewCells[i].length, j = cLen - 1; j > -1; j--) {
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
