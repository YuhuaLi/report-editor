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
  } = { rowStart: 1, columnStart: 1, rowEnd: 1, columnEnd: 1 };
  state: any = {
    isSelectCell: false,
    isScrollYThumbHover: false,
    isSelectScrollYThumb: false,
  };
  isTicking = false;
  scrollLeft = 0;
  scrollTop = 0;
  scrollWidth = 0;
  scrollHeight = 0;
  clientWidth = 0;
  clientHeight = 0;
  mousePoint: any;
  autoScrollTimeoutID: any;

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
    console.log(this.offsetWidth);
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
    this.scrollWidth =
      this.cells[0][this.cells[0].length - 1].x +
      this.cells[0][this.cells[0].length - 1].width;
    this.scrollHeight =
      this.cells[this.cells.length - 1][0].y +
      this.cells[this.cells.length - 1][0].height -
      this.offsetHeight;
    console.log(this.scrollWidth, this.scrollHeight);
    console.log(this.cells);

    this.drawPanel();
    this.setActive(this.activeRange);
  }

  drawScrollBarX() {
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
    this.ctx.fillStyle = this.state.isScrollXThumbHover
      ? Style.scrollBarThumbActiveColor
      : Style.scrollBarThumbColor;

    let scrollXThumbHeight =
      (this.clientWidth / this.scrollWidth) * this.clientWidth;
    if (scrollXThumbHeight < Style.scrollBarThumbMinSize) {
      scrollXThumbHeight = Style.scrollBarThumbMinSize;
    }

    this.roundedRect(
      this.offsetWidth +
        (this.scrollLeft * (this.clientWidth - scrollXThumbHeight)) /
          (this.scrollWidth - this.clientWidth),
      this.height - Style.scrollBarWidth + Style.scrollBarThumbMargin,
      scrollXThumbHeight,
      Style.scrollBarWidth - 2 * Style.scrollBarThumbMargin,
      Style.scrollBarThumbRadius
    );

    this.ctx.restore();
  }
  drawScrollBarY() {
    this.ctx.save();
    this.ctx.fillStyle = Style.scrollBarBackgroundColor;
    this.ctx.strokeStyle = Style.scrollBarBorderColor;
    this.ctx.lineWidth = Style.scrollBarBorderWidth;
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

    this.ctx.fillStyle = this.state.isScrollYThumbHover
      ? Style.scrollBarThumbActiveColor
      : Style.scrollBarThumbColor;

    let scrollYThumbHeight =
      (this.clientHeight / this.scrollHeight) * this.clientHeight;
    if (scrollYThumbHeight < Style.scrollBarThumbMinSize) {
      scrollYThumbHeight = Style.scrollBarThumbMinSize;
    }
    this.roundedRect(
      this.width - Style.scrollBarWidth + Style.scrollBarThumbMargin,
      this.offsetHeight +
        (this.scrollTop * (this.clientHeight - scrollYThumbHeight)) /
          (this.scrollHeight - this.clientHeight),
      Style.scrollBarWidth - 2 * Style.scrollBarThumbMargin,
      scrollYThumbHeight,
      Style.scrollBarThumbRadius
    );
    this.ctx.restore();
  }

  drawScrollBar() {
    this.drawScrollBarX();
    this.drawScrollBarY();
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
    const startRowIndex =
      this.cells.slice(1).findIndex((row, index) => {
        return (
          row[0].y - this.scrollTop <= this.offsetHeight &&
          row[0].y + row[0].height - this.scrollTop >= this.offsetHeight
        );
      }) + 1;
    const startColumnIndex =
      this.cells[0].slice(1).findIndex((cell, index) => {
        return (
          cell.x - this.scrollLeft <= this.offsetWidth &&
          cell.x + cell.width - this.scrollLeft >= this.offsetWidth
        );
      }) + 1;
    console.log(
      this.cells[0].slice(1).findIndex((cell, index) => {
        return (
          cell.x - this.scrollLeft <= this.offsetWidth &&
          cell.x + cell.width - this.scrollLeft >= this.offsetWidth
        );
      })
    );
    console.log(this.clientWidth);
    console.log(
      startRowIndex,
      startColumnIndex,
      this.scrollLeft,
      this.cells[0][4]
    );
    this.viewCells = [
      [
        this.cells[0][0],
        ...this.cells[0].slice(
          startColumnIndex,
          startColumnIndex + this.viewColumnCount - 1
        ),
      ],
      ...this.cells
        .slice(startRowIndex, startRowIndex + this.viewRowCount - 1)
        .map((row) => [
          row[0],
          ...row.slice(
            startColumnIndex,
            startColumnIndex + this.viewColumnCount - 1
          ),
        ]),
    ];
    console.log(this.viewCells);
    console.log(this.scrollLeft)
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
    console.log('mouseover', event);
  }

  onMouseOut(event: MouseEvent) {
    this.state.isSelectCell = false;
    this.state.isScrollXThumbHover = false;
    this.state.isScrollYThumbHover = false;
    this.state.isSelectScrollXThumb = false;
    this.state.isSelectScrollYThumb = false;
    this.drawScrollBar();
    this.mousePoint = null;
    console.log('mouseout', event);
  }

  onMouseEnter(event) {
    console.log('mouseenter', event);
  }

  onMouseLeave(event) {
    console.log('mouseleave', event);
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

  inThumbAreaOfScrollBarX(x: number, y: number, judegedInScrollBarX = false) {
    if (!judegedInScrollBarX && !this.inScrollXBarArea(x, y)) {
      return;
    }
    let scrollXThumbHeight =
      (this.clientWidth / this.scrollWidth) * this.clientWidth;
    if (scrollXThumbHeight < Style.scrollBarThumbMinSize) {
      scrollXThumbHeight = Style.scrollBarThumbMinSize;
    }
    const scrollYThumbLeft =
      this.offsetWidth +
      (this.scrollLeft * (this.clientWidth - scrollXThumbHeight)) /
        (this.scrollWidth - this.clientWidth);
    return inRange(x, scrollYThumbLeft, scrollYThumbLeft + scrollXThumbHeight);
  }

  inThumbAreaOfScrollBarY(x: number, y: number, judegedInScrollBarY = false) {
    if (!judegedInScrollBarY && !this.inScrollYBarArea(x, y)) {
      return false;
    }
    let scrollYThumbHeight =
      (this.clientHeight / this.scrollHeight) * this.clientHeight;
    if (scrollYThumbHeight < Style.scrollBarThumbMinSize) {
      scrollYThumbHeight = Style.scrollBarThumbMinSize;
    }
    const scrollYThumbTop =
      this.offsetHeight +
      (this.scrollTop * (this.clientHeight - scrollYThumbHeight)) /
        (this.scrollHeight - this.clientHeight);
    return inRange(y, scrollYThumbTop, scrollYThumbTop + scrollYThumbHeight);
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
    this.mousePoint = { x: event.clientX, y: event.clientY };
    if (this.inSelectAllArea(event.clientX, event.clientY)) {
      console.log('all');
    } else if (this.inRulerXArea(event.clientX, event.clientY)) {
      console.log('rulerx');
    } else if (this.inRulerYArea(event.clientX, event.clientY)) {
      console.log('rulery');
    } else if (this.inScrollXBarArea(event.clientX, event.clientY)) {
      console.log('scrollx');
      if (this.inThumbAreaOfScrollBarX(event.clientX, event.clientY, true)) {
        this.state.isSelectScrollXThumb = true;
      }
    } else if (this.inScrollYBarArea(event.clientX, event.clientY)) {
      console.log('scrolly');
      if (this.inThumbAreaOfScrollBarY(event.clientX, event.clientY, true)) {
        this.state.isSelectScrollYThumb = true;
      }
    } else if (this.inCellArea(event.clientX, event.clientY)) {
      setTimeout(
        () =>
          this.mousePoint &&
          this.autoScroll(this.mousePoint.x, this.mousePoint.y),
        200
      );
      this.state.isSelectCell = true;
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
            this.setActive({
              rowStart:
                (event.shiftKey &&
                  this.activeRange &&
                  this.activeRange.rowStart) ||
                cell.position.row,
              columnStart:
                (event.shiftKey &&
                  this.activeRange &&
                  this.activeRange.columnStart) ||
                cell.position.column,
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
    if (
      (!this.inCellArea(event.clientX, event.clientY) &&
        !this.state.isSelectCell) ||
      this.state.isSelectScrollYThumb
    ) {
      this.panel.nativeElement.style.cursor = 'default';
    } else {
      this.panel.nativeElement.style.cursor = 'cell';
    }
    const preIsScrollXThumbHover = this.state.isScrollXThumbHover;
    const preIsScrollYThumbHover = this.state.isScrollYThumbHover;
    this.state.isScrollXThumbHover = this.inThumbAreaOfScrollBarX(
      event.clientX,
      event.clientY
    );
    this.state.isScrollYThumbHover = this.inThumbAreaOfScrollBarY(
      event.clientX,
      event.clientY
    );
    if (
      (preIsScrollXThumbHover !== this.state.isScrollXThumbHover &&
        !this.state.isSelectScrollXThumb) ||
      (preIsScrollYThumbHover !== this.state.isScrollYThumbHover &&
        !this.state.isSelectScrollYThumb)
    ) {
      this.drawScrollBar();
    }

    if (!this.isTicking) {
      requestAnimationFrame(() => {
        if (this.state.isSelectCell) {
          this.calcActive(event.clientX, event.clientY);
        } else if (this.state.isSelectScrollYThumb) {
          this.calcScrollY(event.clientX, event.clientY);
        } else if (this.state.isSelectScrollXThumb) {
          this.calcScrollX(event.clientX, event.clientY);
        }
        this.mousePoint = { x: event.clientX, y: event.clientY };
        this.isTicking = false;
      });
    }

    this.isTicking = true;
  }

  autoScroll(x: number, y: number) {
    console.log(x, y, this.mousePoint);
    if (
      x === this.mousePoint.x &&
      y === this.mousePoint.y &&
      this.state.isSelectCell
    ) {
      if (y > this.offsetHeight + this.clientHeight) {
        this.scrollY(Style.cellHeight);
        this.calcActive(this.mousePoint.x, this.mousePoint.y);
      } else if (y < this.offsetHeight) {
        this.scrollY(-1 * Style.cellHeight);
        this.calcActive(this.mousePoint.x, this.mousePoint.y);
      }
      this.autoScrollTimeoutID = setTimeout(
        () =>
          this.mousePoint &&
          this.state.isSelectCell &&
          this.autoScroll(this.mousePoint.x, this.mousePoint.y),
        100
      );
    } else {
      if (this.autoScrollTimeoutID) {
        clearTimeout(this.autoScrollTimeoutID);
        this.autoScrollTimeoutID = null;
      }
      this.autoScrollTimeoutID = setTimeout(
        () =>
          this.mousePoint &&
          this.state.isSelectCell &&
          this.autoScroll(this.mousePoint.x, this.mousePoint.y),
        100
      );
    }
  }

  calcScrollX(x: number, y: number) {
    let scrollXThumbHeight =
      (this.clientWidth / this.scrollWidth) * this.clientWidth;
    if (scrollXThumbHeight < Style.scrollBarThumbMinSize) {
      scrollXThumbHeight = Style.scrollBarThumbMinSize;
    }
    const deltaX =
      ((x - this.mousePoint.x) * (this.scrollWidth - this.clientWidth)) /
      (this.clientWidth - scrollXThumbHeight);
    this.scrollX(deltaX);
  }

  calcScrollY(x: number, y: number) {
    let scrollYThumbHeight =
      (this.clientHeight / this.scrollHeight) * this.clientHeight;
    if (scrollYThumbHeight < Style.scrollBarThumbMinSize) {
      scrollYThumbHeight = Style.scrollBarThumbMinSize;
    }
    const deltaY =
      ((y - this.mousePoint.y) * (this.scrollHeight - this.clientHeight)) /
      (this.clientHeight - scrollYThumbHeight);
    this.scrollY(deltaY);
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
    this.state.isSelectCell = false;
    this.state.isSelectScrollYThumb = false;
    this.state.isSelectScrollXThumb = false;
    this.mousePoint = null;
  }

  scrollX(deltaX: number) {
    if (
      (this.scrollLeft >= this.scrollWidth - this.clientWidth && deltaX > 0) ||
      (this.scrollLeft === 0 && deltaX < 0)
    ) {
      return;
    }
    this.scrollLeft += deltaX;
    if (this.scrollLeft >= this.scrollWidth - this.clientWidth) {
      this.scrollLeft = this.scrollWidth - this.clientWidth;
    } else if (this.scrollLeft <= 0) {
      this.scrollLeft = 0;
    }

    this.drawPanel();
    this.setActive(this.activeRange);
  }

  scrollY(deltaY: number) {
    if (
      (this.scrollTop >= this.scrollHeight - this.clientHeight && deltaY > 0) ||
      (this.scrollTop === 0 && deltaY < 0)
    ) {
      return;
    }
    this.scrollTop += deltaY;
    if (this.scrollTop >= this.scrollHeight - this.clientHeight) {
      this.scrollTop = this.scrollHeight - this.clientHeight;
    } else if (this.scrollTop <= 0) {
      this.scrollTop = 0;
    }

    this.drawPanel();
    this.setActive(this.activeRange);
  }

  onWheel(event: WheelEvent) {
    console.log('wheel', event);
    if (!this.isTicking) {
      requestAnimationFrame(() => {
        this.scrollY(event.deltaY);

        this.isTicking = false;
      });
    }
    this.isTicking = true;
  }

  onDblClick(event: MouseEvent) {
    console.log('dblClick', event);
  }
}
