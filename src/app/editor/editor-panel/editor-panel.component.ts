import { Style } from './style.enum';
import {
  Component,
  OnInit,
  ViewChild,
  ElementRef,
  AfterViewInit,
} from '@angular/core';
import Cell from './cell.calss';
import { inRange } from 'src/app/core/decorator/utils/function';

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
    columnStart: number;
    rowEnd: number;
    columnEnd: number;
  };
  state: any = {};
  isTicking = false;
  scrollLeft = 0;
  scrollTop = 0;
  scrollWidth = 0;
  scrollHeight = 0;
  clientWidth = 0;
  clientHeight = 0;

  ctx: CanvasRenderingContext2D;

  constructor(private elmentRef: ElementRef) {}

  ngOnInit(): void {
    this.width = this.elmentRef.nativeElement.offsetWidth;
    this.height = this.elmentRef.nativeElement.offsetHeight;
    this.clientHeight = this.height - this.offsetHeight - Style.scrollBarWidth;
  }

  ngAfterViewInit() {
    this.ctx = this.panel.nativeElement.getContext('2d');
    this.viewRowCount =
      Math.ceil((this.height - this.offsetHeight) / Style.cellHeight) + 2;
    this.ctx.font = `${Style.rulerCellFontWeight} ${Style.rulerCellFontSize}px ${Style.rulerCellFontFamily}`;
    this.offsetWidth = Math.ceil(
      this.ctx.measureText(`  ${this.viewRowCount}  `).width
    );
    this.clientWidth = this.width - this.offsetWidth - Style.scrollBarWidth;
    this.viewColumnCount =
      Math.ceil((this.width - this.offsetWidth) / Style.cellWidth) + 2;

    this.cells = Array.from({ length: this.viewRowCount + 1 }, (rv, rk) => {
      const isXRuler = rk === 0;
      return Array.from({ length: this.viewColumnCount + 1 }, (cv, ck) => {
        const isYRuler = ck === 0;
        return {
          position: { row: rk, column: ck },
          x: ck === 0 ? 0 : this.offsetWidth + (ck - 1) * Style.cellWidth,
          y: isXRuler ? 0 : this.offsetHeight + (rk - 1) * Style.cellHeight,
          width: isYRuler ? this.offsetWidth : Style.cellWidth,
          height: isXRuler ? this.offsetHeight : Style.cellHeight,
          type:
            isXRuler && isYRuler
              ? 'all'
              : isXRuler || isYRuler
              ? 'ruler'
              : 'cell',
          content: {
            value:
              isYRuler && !isXRuler
                ? this.generateRowNum(rk)
                : isXRuler && !isYRuler
                ? this.generateColumnNum(ck)
                : 'null',
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
    this.scrollHeight =
      this.cells[this.cells.length - 1][0].y +
      this.cells[this.cells.length - 1][0].height -
      this.offsetHeight;
    console.log(this.scrollHeight);
    console.log(this.cells);
    // this.viewCells = this.cells.slice(1).map((cells) => cells.slice(1));
    // this.viewCells = this.cells;
    this.drawPanel();
    this.setActive(this.activeRange);
  }

  drawScrollBar() {
    this.ctx.save();
    this.ctx.fillStyle = Style.scrollBarBackgroundColor;
    this.ctx.strokeStyle = Style.scrollBarBorderColor;
    this.ctx.lineWidth = Style.scrollBarBorderWidth;
    this.ctx.fillRect(
      this.offsetWidth,
      this.height - Style.scrollBarWidth,
      this.width - Style.scrollBarWidth,
      this.height
    );
    this.ctx.strokeRect(
      this.offsetWidth,
      this.height - Style.scrollBarWidth,
      this.width - Style.scrollBarWidth,
      this.height
    );
    this.ctx.fillRect(
      this.width - Style.scrollBarWidth,
      this.offsetHeight,
      this.width,
      this.height - Style.scrollBarWidth
    );
    this.ctx.strokeRect(
      this.width - Style.scrollBarWidth,
      this.offsetHeight,
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
      this.offsetHeight +
        Math.floor(this.scrollTop * (this.clientHeight / this.scrollHeight)),
      Style.scrollBarWidth - 2 * Style.scrollBarThumbMargin,
      Math.floor((this.clientHeight / this.scrollHeight) * this.clientHeight),
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
        columns[i].x - this.scrollLeft,
        columns[i].y,
        columns[i].width,
        columns[i].height
      );
    }
    for (let len = rows.length, i = len - 1; i >= 0; i--) {
      this.ctx.fillRect(
        rows[i].x,
        rows[i].y - this.scrollTop,
        rows[i].width,
        rows[i].height
      );
    }
    this.ctx.fillStyle = columns[0].color;
    this.ctx.strokeStyle = Style.rulerCellBorderColor;
    for (let len = columns.length, i = len - 1; i >= 0; i--) {
      this.ctx.strokeRect(
        columns[i].x - this.scrollLeft,
        columns[i].y,
        columns[i].width,
        columns[i].height
      );
      if (columns[i].content.value) {
        this.ctx.fillText(
          columns[i].content.value,
          columns[i].x - this.scrollLeft + columns[i].width / 2,
          columns[i].y + columns[i].height / 2,
          columns[i].width - 2 * columns[0].borderWidth
        );
      }
    }
    for (let len = rows.length, i = len - 1; i >= 0; i--) {
      this.ctx.strokeRect(
        rows[i].x,
        rows[i].y - this.scrollTop,
        rows[i].width,
        rows[i].height
      );
      if (rows[i].content.value) {
        this.ctx.fillText(
          rows[i].content.value,
          rows[i].x + rows[i].width / 2,
          rows[i].y - this.scrollTop + rows[i].height / 2,
          rows[i].width - 2 * columns[0].borderWidth
        );
      }
    }
    this.ctx.fillStyle = Style.cellBackgroundColor;
    this.ctx.fillRect(0, 0, this.offsetWidth, this.offsetHeight);
    this.ctx.strokeRect(0, 0, this.offsetWidth, this.offsetHeight);
    this.ctx.restore();
  }

  drawPanel() {
    const startIndex =
      this.cells.slice(1).findIndex((row, index) => {
        console.log(index);
        return (
          row[0].y - this.scrollTop <= this.offsetHeight &&
          this.cells[index + 2][0].y - this.scrollTop >= this.offsetHeight
        );
      }) + 1;
    this.viewCells = [
      this.cells[0],
      ...this.cells.slice(startIndex, startIndex + this.viewRowCount - 1),
    ];

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
    const x = cell.x - this.scrollLeft;
    const y = cell.y - this.scrollTop;
    const width = cell.width;
    const height = cell.height;
    this.ctx.clearRect(x, y, width, height);
    if (cell.background) {
      if (this.ctx.fillStyle !== cell.background) {
        this.ctx.fillStyle = cell.background;
      }
      this.ctx.fillRect(x, y, width, height);
    }
    if (this.ctx.strokeStyle !== cell.borderColor) {
      this.ctx.strokeStyle = cell.borderColor;
    }
    this.ctx.strokeRect(x, y, width, height);
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
        x + width / 2,
        y + height / 2,
        width - 2 * cell.borderWidth
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
          // if (
          //   i > rowStart + 1 &&
          //   i < rowEnd - 1 &&
          //   j > columnStart + 1 &&
          //   j < columnEnd - 1
          // ) {
          //   continue;
          // }
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
      this.ctx.strokeStyle = Style.activeCellBorderColor;
      this.ctx.shadowColor = Style.activeCellShadowColor;
      this.ctx.shadowBlur = Style.activeCellShadowBlur;
      this.ctx.strokeRect(
        this.cells[rowStart][columnStart].x - this.scrollLeft,
        this.cells[rowStart][columnStart].y - this.scrollTop,
        this.cells[rowEnd][columnEnd].x -
          this.cells[rowStart][columnStart].x +
          this.cells[rowEnd][columnEnd].width,
        this.cells[rowEnd][columnEnd].y -
          this.cells[rowStart][columnStart].y +
          this.cells[rowEnd][columnEnd].height
      );
      this.ctx.fillStyle = Style.selectedCellBackgroundColor;
      this.ctx.fillRect(
        this.cells[rowStart][columnStart].x - this.scrollLeft,
        this.cells[rowStart][columnStart].y - this.scrollTop,
        this.cells[rowEnd][columnEnd].x -
          this.cells[rowStart][columnStart].x +
          this.cells[rowEnd][columnEnd].width,
        this.cells[rowEnd][columnEnd].y -
          this.cells[rowStart][columnStart].y +
          this.cells[rowEnd][columnEnd].height
      );
    }
    this.ctx.restore();
    this.drawRuler();
    this.drawScrollBar();
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

  inCellArea(x: number, y: number) {
    return (
      inRange(x, this.offsetWidth, this.offsetWidth + this.clientWidth, true) &&
      inRange(y, this.offsetHeight, this.offsetHeight + this.clientHeight, true)
    );
  }

  inRulerXArea(x: number, y: number) {
    return (
      inRange(x, this.offsetWidth, this.width, true) &&
      inRange(y, 0, this.offsetHeight, true)
    );
  }
  inRulerYArea(x: number, y: number) {
    return (
      inRange(x, 0, this.offsetWidth, true) &&
      inRange(y, this.offsetHeight, this.height, true)
    );
  }

  inScrollYBarArea(x: number, y: number) {
    return (
      inRange(x, this.offsetWidth + this.clientWidth, this.width, true) &&
      inRange(y, this.offsetHeight, this.offsetHeight + this.clientHeight, true)
    );
  }
  inScrollXBarArea(x: number, y: number) {
    return (
      inRange(x, this.offsetWidth, this.width, true) &&
      inRange(y, this.offsetHeight + this.clientHeight, this.height, true)
    );
  }
  inSelectAllArea(x: number, y: number) {
    return (
      inRange(x, 0, this.offsetWidth, true) &&
      inRange(y, 0, this.offsetHeight, true)
    );
  }

  onMouseDown(event: MouseEvent) {
    if (event.button === 2) {
      event.returnValue = false;
      return;
    }
    if (this.inSelectAllArea(event.clientX, event.clientY)) {
      console.log('all');
    } else if (this.inRulerXArea(event.clientX, event.clientY)) {
      console.log('rulerx');
    } else if (this.inRulerYArea(event.clientX, event.clientY)) {
      console.log('rulery');
    } else if (this.inScrollXBarArea(event.clientX, event.clientY)) {
      console.log('scrollx');
    } else if (this.inScrollYBarArea(event.clientX, event.clientY)) {
      console.log('scrolly');
    } else if (this.inCellArea(event.clientX, event.clientY)) {
      this.state.isMouseDown = true;
      const canvas = document.createElement('canvas');
      canvas.width = this.width;
      canvas.height = this.height;
      const ctx = canvas.getContext('2d');
      for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
        for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
          const cell = this.viewCells[i][j];
          ctx.clearRect(0, 0, this.width, this.height);
          ctx.rect(
            cell.x - this.scrollLeft,
            cell.y - this.scrollTop,
            cell.width,
            cell.height
          );
          if (ctx.isPointInPath(event.clientX, event.clientY)) {
            console.log(cell.position);
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
  }

  // @throttle(20)
  onMouseMove(event: MouseEvent) {
    if (!this.inCellArea(event.clientX, event.clientY)) {
      this.panel.nativeElement.style.cursor = 'default';
    } else {
      this.panel.nativeElement.style.cursor = 'cell';
    }
    if (this.state.isMouseDown) {
      if (!this.isTicking) {
        requestAnimationFrame(() => {
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
        ctx.rect(
          cell.x - this.scrollLeft,
          cell.y - this.scrollTop,
          cell.width,
          cell.height
        );
        if (ctx.isPointInPath(x, y)) {
          console.log(cell.position);
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

  onWheel(event: WheelEvent) {
    console.log('wheel', event);
    if (event.deltaY) {
      if (
        (this.scrollTop >= this.scrollHeight - this.clientHeight &&
          event.deltaY > 0) ||
        (this.scrollTop === 0 && event.deltaY < 0)
      ) {
        return;
      }
      this.scrollTop += event.deltaY;
      if (this.scrollTop >= this.scrollHeight - this.clientHeight) {
        this.scrollTop = this.scrollHeight - this.clientHeight;
      } else if (this.scrollTop <= 0) {
        this.scrollTop = 0;
      }
      this.drawPanel();
      this.setActive(this.activeRange);
    }
  }

  onDblClick(event: MouseEvent) {
    console.log('dblClick', event);
  }
}
