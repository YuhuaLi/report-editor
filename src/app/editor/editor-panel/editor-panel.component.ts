import { Style } from '../../core/model/style.const';
import {
  Component,
  OnInit,
  ViewChild,
  ElementRef,
  AfterViewInit,
  HostListener,
} from '@angular/core';
import Cell from './cell.calss';
import { inRange } from 'src/app/core/decorator/utils/function';
import { CellRange } from 'src/app/core/model/cell-range.class';
import { KeyCode } from 'src/app/core/model/key-code.enmu';

@Component({
  selector: 'app-editor-panel',
  templateUrl: './editor-panel.component.html',
  styleUrls: ['./editor-panel.component.scss'],
})
export class EditorPanelComponent implements OnInit, AfterViewInit {
  @ViewChild('panel') panel: ElementRef;
  @ViewChild('actionPanel') actionPanel: ElementRef;

  width = 0;
  height = 0;
  viewRowCount = 0;
  viewColumnCount = 0;
  offsetLeft = 0;
  offsetTop = 30;
  columns = [];
  rows = [];
  cells: Cell[][] = [];
  viewCells: Cell[][] = [];
  editingCell: Cell;
  scrollBarWidth = Style.scrollBarWidth;
  // activeRange: {
  //   rowStart: number;
  //   columnStart: number;
  //   rowEnd: number;
  //   columnEnd: number;
  // } = { rowStart: 1, columnStart: 1, rowEnd: 1, columnEnd: 1 };
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
  activeCellPos: any = { row: 1, column: 1, rangeIndex: 0 };
  activeArr: CellRange[] = [
    { rowStart: 1, columnStart: 1, rowEnd: 1, columnEnd: 1 },
  ];
  unActiveRange: CellRange;
  resizeColumnCell: Cell;
  resizeRowCell: Cell;

  ctx: CanvasRenderingContext2D;
  actionCtx: CanvasRenderingContext2D;

  constructor(private elmentRef: ElementRef) {}

  @HostListener('window:resize', ['$event'])
  onResize(event) {
    console.log('resize');
    this.width = this.elmentRef.nativeElement.offsetWidth;
    this.height = this.elmentRef.nativeElement.offsetHeight;
    this.clientWidth = this.width - this.offsetLeft - Style.scrollBarWidth;
    this.clientHeight = this.height - this.offsetTop - Style.scrollBarWidth;
    this.panel.nativeElement.width = this.width;
    this.panel.nativeElement.height = this.height;
    this.actionPanel.nativeElement.width = this.width;
    this.actionPanel.nativeElement.height = this.height;

    this.refreshView();
  }

  ngOnInit(): void {
    this.width = this.elmentRef.nativeElement.offsetWidth;
    this.height = this.elmentRef.nativeElement.offsetHeight;

    this.clientHeight = this.height - this.offsetTop - Style.scrollBarWidth;
  }

  ngAfterViewInit() {
    this.ctx = this.panel.nativeElement.getContext('2d');
    this.actionCtx = this.actionPanel.nativeElement.getContext('2d');
    this.viewRowCount =
      Math.ceil((this.height - this.offsetTop) / Style.cellHeight) + 2;
    this.ctx.font = `${Style.rulerCellFontWeight} ${Style.rulerCellFontSize}px ${Style.rulerCellFontFamily}`;
    this.offsetLeft = Math.ceil(
      this.ctx.measureText(`     ${this.viewRowCount}    `).width
    );
    this.clientWidth = this.width - this.offsetLeft - Style.scrollBarWidth;
    this.viewColumnCount =
      Math.ceil((this.width - this.offsetLeft) / Style.cellWidth) + 2;

    this.columns = Array.from({ length: this.viewColumnCount + 1 }, (v, k) => ({
      x: k === 0 ? 0 : this.offsetLeft + (k - 1) * Style.cellWidth,
      width: k === 0 ? this.offsetLeft : Style.cellWidth,
    }));
    this.rows = Array.from({ length: this.viewRowCount + 1 }, (v, k) => ({
      y: k === 0 ? 0 : this.offsetTop + (k - 1) * Style.cellHeight,
      height: k === 0 ? this.offsetTop : Style.cellHeight,
    }));

    this.cells = Array.from({ length: this.viewRowCount + 1 }, (rv, rk) => {
      return Array.from({ length: this.viewColumnCount + 1 }, (cv, ck) =>
        this.createCell(rk, ck)
      ).reduce((acc, cur) => acc.push(cur) && acc, []);
    });
    this.scrollWidth =
      this.cells[0][this.cells[0].length - 1].x +
      this.cells[0][this.cells[0].length - 1].width -
      this.offsetLeft;
    this.scrollHeight =
      this.cells[this.cells.length - 1][0].y +
      this.cells[this.cells.length - 1][0].height -
      this.offsetTop;

    this.panel.nativeElement.width = this.width;
    this.panel.nativeElement.height = this.height;
    this.actionPanel.nativeElement.width = this.width;
    this.actionPanel.nativeElement.height = this.height;

    this.panel.nativeElement.focus();

    this.refreshView();
  }

  refreshView() {
    // let startRowIndex = this.cells.slice(1).findIndex((row, index) => {
    //   return (
    //     row[0].y - this.scrollTop <= this.offsetTop &&
    //     row[0].y + row[0].height - this.scrollTop >= this.offsetTop
    //   );
    // });
    let startRowIndex = this.rows
      .slice(1)
      .findIndex(
        (row, index) =>
          row.y - this.scrollTop <= this.offsetTop &&
          row.y + row.height - this.scrollTop >= this.offsetTop
      );
    startRowIndex = startRowIndex ? startRowIndex : startRowIndex + 1;
    // let startColumnIndex = this.cells[0].slice(1).findIndex((cell, index) => {
    //   return (
    //     cell.x - this.scrollLeft <= this.offsetLeft &&
    //     cell.x + cell.width - this.scrollLeft >= this.offsetLeft
    //   );
    // });
    let startColumnIndex = this.columns
      .slice(1)
      .findIndex(
        (col, index) =>
          col.x - this.scrollLeft <= this.offsetLeft &&
          col.x + col.width - this.scrollLeft >= this.offsetLeft
      );
    startColumnIndex = startColumnIndex
      ? startColumnIndex
      : startColumnIndex + 1;
    this.viewCells = [
      [
        this.cells[0][0],
        ...this.cells[0].slice(
          startColumnIndex,
          startColumnIndex + this.viewColumnCount
        ),
      ],
      ...this.cells
        .slice(startRowIndex, startRowIndex + this.viewRowCount)
        .map((row) => [
          row[0],
          ...row.slice(
            startColumnIndex,
            startColumnIndex + this.viewColumnCount
          ),
        ]),
    ];

    const canvas = document.createElement('canvas');
    canvas.width = this.width;
    canvas.height = this.height;
    const ctx = canvas.getContext('2d');
    this.setActive();
    this.drawPanel(ctx);
    this.drawRuler(ctx);
    this.drawScrollBar(ctx);
    this.ctx.drawImage(canvas, 0, 0);
  }

  createCell(rk: number, ck: number): Cell {
    const isXRuler = rk === 0;
    const isYRuler = ck === 0;
    return {
      // cells: this.cells,
      // offsetLeft: this.offsetLeft,
      // offsetTop: this.offsetTop,
      columns: this.columns,
      rows: this.rows,
      position: { row: rk, column: ck },
      // x: ck === 0 ? 0 : this.offsetLeft + (ck - 1) * Style.cellWidth,
      get x() {
        if (!this.columns[this.position.column]) {
          this.columns[this.position.column] = {
            x:
              this.columns[this.position.column - 1].x +
              this.columns[this.position.column - 1].width,
          };
        } else if (!('x' in this.columns[this.position.column])) {
          this.columns[this.position.column].x =
            this.columns[this.position.column - 1].x +
            this.columns[this.position.column - 1].width;
        }
        return this.columns[this.position.column].x;
      },
      set x(val) {
        this.columns[this.position.column].x = val;
      },
      get y() {
        if (!this.rows[this.position.row]) {
          this.rows[this.position.row] = {
            y:
              this.rows[this.position.row - 1].y +
              this.rows[this.position.row - 1].height,
          };
        } else if (!('y' in this.rows[this.position.row])) {
          this.rows[this.position.row].y =
            this.rows[this.position.row - 1].y +
            this.rows[this.position.row - 1].height;
        }
        return this.rows[this.position.row].y;
      },
      set y(val) {
        this.rows[this.position.row].y = val;
      },
      // y: isXRuler ? 0 : this.offsetTop + (rk - 1) * Style.cellHeight,
      // width: isYRuler ? this.offsetLeft : Style.cellWidth,
      get width() {
        if (!this.columns[this.position.column]) {
          this.columns[this.position.column] = { width: Style.cellWidth };
        } else if (!('width' in this.columns[this.position.column])) {
          this.columns[this.position.column].width = Style.cellWidth;
        }
        return this.columns[this.position.column].width;
      },
      set width(val) {
        this.columns[this.position.column].width = val;
      },
      // height: isXRuler ? this.offsetTop : Style.cellHeight,
      get height() {
        if (!this.rows[this.position.row]) {
          this.rows[this.position.row] = { height: Style.cellHeight };
        } else if (!('height' in this.rows[this.position.row])) {
          this.rows[this.position.row].width = Style.cellHeight;
        }
        return this.rows[this.position.row].height;
      },
      set height(val) {
        this.rows[this.position.row].height = val;
      },
      type:
        isXRuler && isYRuler ? 'all' : isXRuler || isYRuler ? 'ruler' : 'cell',
      content: {
        value:
          isYRuler && !isXRuler
            ? this.generateRowNum(rk)
            : isXRuler && !isYRuler
            ? this.generateColumnNum(ck)
            : null,
        previousValue: null,
      },
      fontWeight:
        isXRuler || isYRuler ? Style.rulerCellFontWeight : Style.cellFontWeight,
      textAlign: Style.cellTextAlign as CanvasTextAlign,
      textBaseline: Style.cellTextBaseline as CanvasTextBaseline,
      fontStyle: Style.cellFontStyle,
      fontFamily:
        isXRuler || isYRuler ? Style.rulerCellFontFamily : Style.cellFontFamily,
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
  }

  drawScrollBarX(ctx: CanvasRenderingContext2D) {
    ctx.save();
    ctx.fillStyle = Style.scrollBarBackgroundColor;
    ctx.strokeStyle = Style.scrollBarBorderColor;
    ctx.lineWidth = Style.scrollBarBorderWidth;
    ctx.fillRect(
      this.offsetLeft,
      this.height - Style.scrollBarWidth,
      this.width - Style.scrollBarWidth,
      this.height
    );
    ctx.strokeRect(
      this.offsetLeft,
      this.height - Style.scrollBarWidth,
      this.width - Style.scrollBarWidth,
      this.height
    );
    ctx.fillRect(
      this.width - Style.scrollBarWidth,
      this.offsetTop,
      this.width,
      this.height - Style.scrollBarWidth
    );
    ctx.fillStyle =
      this.state.isScrollXThumbHover || this.state.isSelectScrollXThumb
        ? Style.scrollBarThumbActiveColor
        : Style.scrollBarThumbColor;

    const scrollXThumbHeight = this.getScrollXThumbHeight();

    this.roundedRect(
      ctx,
      this.getScrollXThumbLeft(scrollXThumbHeight),
      this.height - Style.scrollBarWidth + Style.scrollBarThumbMargin,
      scrollXThumbHeight,
      Style.scrollBarWidth - 2 * Style.scrollBarThumbMargin,
      Style.scrollBarThumbRadius
    );

    ctx.restore();
  }
  drawScrollBarY(ctx: CanvasRenderingContext2D) {
    ctx.save();
    ctx.fillStyle = Style.scrollBarBackgroundColor;
    ctx.strokeStyle = Style.scrollBarBorderColor;
    ctx.lineWidth = Style.scrollBarBorderWidth;
    ctx.strokeRect(
      this.width - Style.scrollBarWidth,
      this.offsetTop,
      this.width,
      this.height - Style.scrollBarWidth
    );
    ctx.fillRect(
      this.width - Style.scrollBarWidth,
      this.height - Style.scrollBarWidth,
      this.width,
      this.height
    );

    ctx.fillStyle =
      this.state.isScrollYThumbHover || this.state.isSelectScrollYThumb
        ? Style.scrollBarThumbActiveColor
        : Style.scrollBarThumbColor;

    const scrollYThumbHeight = this.getScrollYThumbHeight();
    this.roundedRect(
      ctx,
      this.width - Style.scrollBarWidth + Style.scrollBarThumbMargin,
      this.getScrollYThumbTop(scrollYThumbHeight),
      Style.scrollBarWidth - 2 * Style.scrollBarThumbMargin,
      scrollYThumbHeight,
      Style.scrollBarThumbRadius
    );
    ctx.restore();
  }

  drawScrollBar(ctx: CanvasRenderingContext2D) {
    this.drawScrollBarX(ctx);
    this.drawScrollBarY(ctx);
  }

  roundedRect(
    ctx: CanvasRenderingContext2D,
    x: number,
    y: number,
    width: number,
    height: number,
    radius: number
  ) {
    ctx.beginPath();
    ctx.moveTo(x, y + radius);
    ctx.lineTo(x, y + height - radius);
    ctx.quadraticCurveTo(x, y + height, x + radius, y + height);
    ctx.lineTo(x + width - radius, y + height);
    ctx.quadraticCurveTo(x + width, y + height, x + width, y + height - radius);
    ctx.lineTo(x + width, y + radius);
    ctx.quadraticCurveTo(x + width, y, x + width - radius, y);
    ctx.lineTo(x + radius, y);
    ctx.quadraticCurveTo(x, y, x, y + radius);
    ctx.fill();
  }

  drawRuler(ctx: CanvasRenderingContext2D) {
    ctx.save();
    const columns = this.viewCells[0].slice(1);
    const rows = this.viewCells.slice(1).map((cells) => cells[0]);
    ctx.textAlign = columns[0].textAlign as CanvasTextAlign;
    ctx.textBaseline = columns[0].textBaseline as CanvasTextBaseline;
    if (
      ctx.font !==
      `${columns[0].fontStyle} ${columns[0].fontWeight} ${columns[0].fontSize}px ${columns[0].fontFamily}`
    ) {
      ctx.font = `${columns[0].fontStyle} ${columns[0].fontWeight} ${columns[0].fontSize}px ${columns[0].fontFamily}`;
    }
    for (let len = columns.length, i = len - 1; i >= 0; i--) {
      if (
        this.activeArr.length &&
        this.activeArr.find((range) =>
          inRange(
            columns[i].position.column,
            range.columnStart,
            range.columnEnd,
            true
          )
        )
      ) {
        ctx.fillStyle = Style.activeRulerCellBacgroundColor;
      } else if (ctx.fillStyle !== Style.rulerCellBackgroundColor) {
        ctx.fillStyle = Style.rulerCellBackgroundColor;
      }
      ctx.fillRect(
        columns[i].x - this.scrollLeft,
        columns[i].y,
        columns[i].width,
        columns[i].height
      );
    }
    for (let len = rows.length, i = len - 1; i >= 0; i--) {
      if (
        this.activeArr.length &&
        this.activeArr.find((range) =>
          inRange(rows[i].position.row, range.rowStart, range.rowEnd, true)
        )
      ) {
        ctx.fillStyle = Style.activeRulerCellBacgroundColor;
      } else if (ctx.fillStyle !== Style.rulerCellBackgroundColor) {
        ctx.fillStyle = Style.rulerCellBackgroundColor;
      }
      ctx.fillRect(
        rows[i].x,
        rows[i].y - this.scrollTop,
        rows[i].width,
        rows[i].height
      );
    }
    ctx.fillStyle = columns[0].color;
    ctx.strokeStyle = Style.rulerCellBorderColor;
    for (let len = columns.length, i = len - 1; i >= 0; i--) {
      ctx.strokeRect(
        columns[i].x - this.scrollLeft,
        columns[i].y,
        columns[i].width,
        columns[i].height
      );
      if (columns[i].content.value) {
        ctx.fillText(
          columns[i].content.value,
          columns[i].x - this.scrollLeft + columns[i].width / 2,
          columns[i].y + columns[i].height / 2,
          columns[i].width - 2 * columns[0].borderWidth
        );
      }
    }
    for (let len = rows.length, i = len - 1; i >= 0; i--) {
      ctx.strokeRect(
        rows[i].x,
        rows[i].y - this.scrollTop,
        rows[i].width,
        rows[i].height
      );
      if (rows[i].content.value) {
        ctx.fillText(
          rows[i].content.value,
          rows[i].x + rows[i].width / 2,
          rows[i].y - this.scrollTop + rows[i].height / 2,
          rows[i].width - 2 * columns[0].borderWidth
        );
      }
    }
    ctx.fillStyle = Style.cellBackgroundColor;
    ctx.fillRect(0, 0, this.offsetLeft, this.offsetTop);
    ctx.strokeRect(0, 0, this.offsetLeft, this.offsetTop);
    ctx.restore();
  }

  drawPanel(ctx: CanvasRenderingContext2D) {
    for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
      for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
        this.drawCell(ctx, this.viewCells[i][j]);
      }
    }
  }

  drawCell(ctx: CanvasRenderingContext2D, cell: Cell) {
    const x = cell.x - this.scrollLeft;
    const y = cell.y - this.scrollTop;
    const width = cell.width;
    const height = cell.height;
    if (cell.background) {
      if (ctx.fillStyle !== cell.background) {
        ctx.fillStyle = cell.background;
      }
      ctx.fillRect(x, y, width, height);
    }
    if (ctx.strokeStyle !== cell.borderColor) {
      ctx.strokeStyle = cell.borderColor;
    }
    ctx.strokeRect(x, y, width, height);
    if (cell.content.value) {
      if (ctx.fillStyle !== cell.color) {
        ctx.fillStyle = cell.color;
      }
      if (cell.textAlign && ctx.textAlign !== cell.textAlign) {
        ctx.textAlign = cell.textAlign as CanvasTextAlign;
      }
      if (cell.textBaseline && ctx.textBaseline !== cell.textBaseline) {
        ctx.textBaseline = cell.textBaseline as CanvasTextBaseline;
      }
      if (
        ctx.font !==
        `${cell.fontStyle} ${cell.fontWeight} ${cell.fontSize}px ${cell.fontFamily}`
      ) {
        ctx.font = `${cell.fontStyle} ${cell.fontWeight} ${cell.fontSize}px ${cell.fontFamily}`;
      }
      ctx.fillText(
        cell.content.value,
        x + width / 2,
        y + height / 2,
        width - 2 * cell.borderWidth
      );
    }
  }

  setActive() {
    this.actionCtx.clearRect(0, 0, this.width, this.height);

    const canvas = document.createElement('canvas');
    canvas.width = this.width;
    canvas.height = this.height;
    const ctx = canvas.getContext('2d');
    ctx.save();
    if (this.activeArr.length > 1) {
      for (let i = 0, len = this.activeArr.length; i < len; i++) {
        const rowStart = Math.max(
          Math.min(this.activeArr[i].rowEnd, this.activeArr[i].rowStart),
          this.viewCells[1][0].position.row
        );
        const rowEnd = Math.min(
          Math.max(this.activeArr[i].rowEnd, this.activeArr[i].rowStart),
          this.viewCells[this.viewCells.length - 1][0].position.row
        );
        const columnStart = Math.max(
          Math.min(this.activeArr[i].columnStart, this.activeArr[i].columnEnd),
          this.viewCells[0][1].position.column
        );
        const columnEnd = Math.min(
          Math.max(this.activeArr[i].columnStart, this.activeArr[i].columnEnd),
          this.viewCells[0][this.viewCells[0].length - 1].position.column
        );
        ctx.fillStyle = Style.selectedCellBackgroundColor;
        ctx.fillRect(
          this.cells[rowStart][columnStart].x -
            this.scrollLeft +
            3 * Style.cellBorderWidth,
          this.cells[rowStart][columnStart].y -
            this.scrollTop +
            3 * Style.cellBorderWidth,
          this.cells[rowEnd][columnEnd].x -
            this.cells[rowStart][columnStart].x +
            this.cells[rowEnd][columnEnd].width -
            6 * Style.cellBorderWidth,
          this.cells[rowEnd][columnEnd].y -
            this.cells[rowStart][columnStart].y +
            this.cells[rowEnd][columnEnd].height -
            6 * Style.cellBorderWidth
        );
        if (this.activeCellPos) {
          ctx.restore();
          ctx.strokeStyle = Style.activeCellBorderColor;
          ctx.lineWidth = Style.cellBorderWidth;
          ctx.clearRect(
            this.cells[this.activeCellPos.row][this.activeCellPos.column].x -
              this.scrollLeft,
            this.cells[this.activeCellPos.row][this.activeCellPos.column].y -
              this.scrollTop,
            this.cells[this.activeCellPos.row][this.activeCellPos.column].x -
              this.cells[this.activeCellPos.row][this.activeCellPos.column].x +
              this.cells[this.activeCellPos.row][this.activeCellPos.column]
                .width,
            this.cells[this.activeCellPos.row][this.activeCellPos.column].y -
              this.cells[this.activeCellPos.row][this.activeCellPos.column].y +
              this.cells[this.activeCellPos.row][this.activeCellPos.column]
                .height
          );
          ctx.strokeRect(
            this.cells[this.activeCellPos.row][this.activeCellPos.column].x -
              this.scrollLeft +
              3 * Style.cellBorderWidth,
            this.cells[this.activeCellPos.row][this.activeCellPos.column].y -
              this.scrollTop +
              3 * Style.cellBorderWidth,
            this.cells[this.activeCellPos.row][this.activeCellPos.column].x -
              this.cells[this.activeCellPos.row][this.activeCellPos.column].x +
              this.cells[this.activeCellPos.row][this.activeCellPos.column]
                .width -
              6 * Style.cellBorderWidth,
            this.cells[this.activeCellPos.row][this.activeCellPos.column].y -
              this.cells[this.activeCellPos.row][this.activeCellPos.column].y +
              this.cells[this.activeCellPos.row][this.activeCellPos.column]
                .height -
              6 * Style.cellBorderWidth
          );
        }
      }
    } else {
      const rowStart = Math.max(
        Math.min(this.activeArr[0].rowEnd, this.activeArr[0].rowStart),
        this.viewCells[1][0].position.row
      );
      const rowEnd = Math.min(
        Math.max(this.activeArr[0].rowEnd, this.activeArr[0].rowStart),
        this.viewCells[this.viewCells.length - 1][0].position.row
      );
      const columnStart = Math.max(
        Math.min(this.activeArr[0].columnStart, this.activeArr[0].columnEnd),
        this.viewCells[0][1].position.column
      );
      const columnEnd = Math.min(
        Math.max(this.activeArr[0].columnStart, this.activeArr[0].columnEnd),
        this.viewCells[0][this.viewCells[0].length - 1].position.column
      );

      ctx.fillStyle = Style.selectedCellBackgroundColor;
      ctx.fillRect(
        this.cells[rowStart][columnStart].x - this.scrollLeft,
        this.cells[rowStart][columnStart].y - this.scrollTop,
        this.cells[rowEnd][columnEnd].x -
          this.cells[rowStart][columnStart].x +
          this.cells[rowEnd][columnEnd].width,
        this.cells[rowEnd][columnEnd].y -
          this.cells[rowStart][columnStart].y +
          this.cells[rowEnd][columnEnd].height
      );
      if (this.activeCellPos) {
        ctx.clearRect(
          this.cells[this.activeCellPos.row][this.activeCellPos.column].x -
            this.scrollLeft,
          this.cells[this.activeCellPos.row][this.activeCellPos.column].y -
            this.scrollTop,
          this.cells[this.activeCellPos.row][this.activeCellPos.column].x -
            this.cells[this.activeCellPos.row][this.activeCellPos.column].x +
            this.cells[this.activeCellPos.row][this.activeCellPos.column].width,
          this.cells[this.activeCellPos.row][this.activeCellPos.column].y -
            this.cells[this.activeCellPos.row][this.activeCellPos.column].y +
            this.cells[this.activeCellPos.row][this.activeCellPos.column].height
        );
      }
      ctx.strokeStyle = Style.activeCellBorderColor;
      ctx.shadowColor = Style.activeCellShadowColor;
      ctx.shadowBlur = Style.activeCellShadowBlur;
      ctx.strokeRect(
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
    if (this.unActiveRange) {
      ctx.restore();
      const rowStart = Math.max(
        Math.min(this.unActiveRange.rowEnd, this.unActiveRange.rowStart),
        this.viewCells[1][0].position.row
      );
      const rowEnd = Math.min(
        Math.max(this.unActiveRange.rowEnd, this.unActiveRange.rowStart),
        this.viewCells[this.viewCells.length - 1][0].position.row
      );
      const columnStart = Math.max(
        Math.min(this.unActiveRange.columnStart, this.unActiveRange.columnEnd),
        this.viewCells[0][1].position.column
      );
      const columnEnd = Math.min(
        Math.max(this.unActiveRange.columnStart, this.unActiveRange.columnEnd),
        this.viewCells[0][this.viewCells[0].length - 1].position.column
      );
      ctx.strokeStyle = Style.cellBorderColor;
      ctx.lineWidth = 3 * Style.cellBorderWidth;
      ctx.fillStyle = Style.unSelectedCellBackgroundColor;
      ctx.rect(
        this.cells[rowStart][columnStart].x - this.scrollLeft,
        this.cells[rowStart][columnStart].y - this.scrollTop,
        this.cells[rowEnd][columnEnd].x -
          this.cells[rowStart][columnStart].x +
          this.cells[rowEnd][columnEnd].width,
        this.cells[rowEnd][columnEnd].y -
          this.cells[rowStart][columnStart].y +
          this.cells[rowEnd][columnEnd].height
      );
      ctx.fill();
      ctx.stroke();
    }
    this.actionCtx.drawImage(
      canvas,
      this.offsetLeft,
      this.offsetTop,
      this.clientWidth,
      this.clientHeight,
      this.offsetLeft,
      this.offsetTop,
      this.clientWidth,
      this.clientHeight
    );
  }

  // setActive(activeRange?: {
  //   rowStart: number;
  //   rowEnd: number;
  //   columnStart: number;
  //   columnEnd: number;
  // }) {
  //   this.actionCtx.clearRect(0, 0, this.width, this.height);
  //   this.activeRange = activeRange;
  //   if (activeRange) {
  //     const canvas = document.createElement('canvas');
  //     canvas.width = this.width;
  //     canvas.height = this.height;
  //     const ctx = canvas.getContext('2d');

  //     const rowStart = Math.max(
  //       Math.min(this.activeRange.rowEnd, this.activeRange.rowStart),
  //       this.viewCells[1][0].position.row
  //     );
  //     const rowEnd = Math.min(
  //       Math.max(this.activeRange.rowEnd, this.activeRange.rowStart),
  //       this.viewCells[this.viewCells.length - 1][0].position.row
  //     );
  //     const columnStart = Math.max(
  //       Math.min(this.activeRange.columnStart, this.activeRange.columnEnd),
  //       this.viewCells[0][1].position.column
  //     );
  //     const columnEnd = Math.min(
  //       Math.max(this.activeRange.columnStart, this.activeRange.columnEnd),
  //       this.viewCells[0][this.viewCells[0].length - 1].position.column
  //     );
  //     ctx.strokeStyle = Style.activeCellBorderColor;
  //     ctx.shadowColor = Style.activeCellShadowColor;
  //     ctx.shadowBlur = Style.activeCellShadowBlur;
  //     ctx.strokeRect(
  //       this.cells[rowStart][columnStart].x - this.scrollLeft,
  //       this.cells[rowStart][columnStart].y - this.scrollTop,
  //       this.cells[rowEnd][columnEnd].x -
  //         this.cells[rowStart][columnStart].x +
  //         this.cells[rowEnd][columnEnd].width,
  //       this.cells[rowEnd][columnEnd].y -
  //         this.cells[rowStart][columnStart].y +
  //         this.cells[rowEnd][columnEnd].height
  //     );
  //     ctx.fillStyle = Style.selectedCellBackgroundColor;
  //     ctx.fillRect(
  //       this.cells[rowStart][columnStart].x - this.scrollLeft,
  //       this.cells[rowStart][columnStart].y - this.scrollTop,
  //       this.cells[rowEnd][columnEnd].x -
  //         this.cells[rowStart][columnStart].x +
  //         this.cells[rowEnd][columnEnd].width,
  //       this.cells[rowEnd][columnEnd].y -
  //         this.cells[rowStart][columnStart].y +
  //         this.cells[rowEnd][columnEnd].height
  //     );

  //     this.actionCtx.drawImage(
  //       canvas,
  //       this.offsetLeft,
  //       this.offsetTop,
  //       this.clientWidth,
  //       this.clientHeight,
  //       this.offsetLeft,
  //       this.offsetTop,
  //       this.clientWidth,
  //       this.clientHeight
  //     );
  //   }
  // }

  // setActive(activeRange?: {
  //   rowStart: number;
  //   rowEnd: number;
  //   columnStart: number;
  //   columnEnd: number;
  // }) {
  //   this.ctx.save();
  //   if (this.activeRange) {
  //     const rowStart = Math.max(
  //       Math.min(this.activeRange.rowEnd, this.activeRange.rowStart),
  //       this.viewCells[1][0].position.row
  //     );
  //     const rowEnd = Math.min(
  //       Math.max(this.activeRange.rowEnd, this.activeRange.rowStart),
  //       this.viewCells[this.viewCells.length - 1][0].position.row
  //     );
  //     const columnStart = Math.max(
  //       Math.min(this.activeRange.columnStart, this.activeRange.columnEnd),
  //       this.viewCells[0][1].position.column
  //     );
  //     const columnEnd = Math.min(
  //       Math.max(this.activeRange.columnStart, this.activeRange.columnEnd),
  //       this.viewCells[0][this.viewCells[0].length - 1].position.column
  //     );
  //     for (let i = rowStart - 1, rLen = rowEnd + 2; i < rLen; i++) {
  //       for (let j = columnStart - 1, cLen = columnEnd + 2; j < cLen; j++) {
  //         // if (
  //         //   i > rowStart + 1 &&
  //         //   i < rowEnd - 1 &&
  //         //   j > columnStart + 1 &&
  //         //   j < columnEnd - 1
  //         // ) {
  //         //   continue;
  //         // }
  //         if (this.cells[i] && this.cells[i][j]) {
  //           this.drawCell(this.cells[i][j]);
  //         }
  //       }
  //     }
  //   }
  //   this.activeRange = activeRange;
  //   if (activeRange) {
  //     const rowStart = Math.max(
  //       Math.min(this.activeRange.rowEnd, this.activeRange.rowStart),
  //       this.viewCells[1][0].position.row
  //     );
  //     const rowEnd = Math.min(
  //       Math.max(this.activeRange.rowEnd, this.activeRange.rowStart),
  //       this.viewCells[this.viewCells.length - 1][0].position.row
  //     );
  //     const columnStart = Math.max(
  //       Math.min(this.activeRange.columnStart, this.activeRange.columnEnd),
  //       this.viewCells[0][1].position.column
  //     );
  //     const columnEnd = Math.min(
  //       Math.max(this.activeRange.columnStart, this.activeRange.columnEnd),
  //       this.viewCells[0][this.viewCells[0].length - 1].position.column
  //     );
  //     this.ctx.strokeStyle = Style.activeCellBorderColor;
  //     this.ctx.shadowColor = Style.activeCellShadowColor;
  //     this.ctx.shadowBlur = Style.activeCellShadowBlur;
  //     this.ctx.strokeRect(
  //       this.cells[rowStart][columnStart].x - this.scrollLeft,
  //       this.cells[rowStart][columnStart].y - this.scrollTop,
  //       this.cells[rowEnd][columnEnd].x -
  //         this.cells[rowStart][columnStart].x +
  //         this.cells[rowEnd][columnEnd].width,
  //       this.cells[rowEnd][columnEnd].y -
  //         this.cells[rowStart][columnStart].y +
  //         this.cells[rowEnd][columnEnd].height
  //     );
  //     this.ctx.fillStyle = Style.selectedCellBackgroundColor;
  //     this.ctx.fillRect(
  //       this.cells[rowStart][columnStart].x - this.scrollLeft,
  //       this.cells[rowStart][columnStart].y - this.scrollTop,
  //       this.cells[rowEnd][columnEnd].x -
  //         this.cells[rowStart][columnStart].x +
  //         this.cells[rowEnd][columnEnd].width,
  //       this.cells[rowEnd][columnEnd].y -
  //         this.cells[rowStart][columnStart].y +
  //         this.cells[rowEnd][columnEnd].height
  //     );
  //   }
  //   this.ctx.restore();
  //   this.drawRuler();
  //   this.drawScrollBar();
  // }

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
    // console.log('mouseover', event);
  }

  onMouseOut(event: MouseEvent) {
    // console.log('mouseout', event);
    this.state.isSelectCell = false;
    this.state.isScrollXThumbHover = false;
    this.state.isScrollYThumbHover = false;
    this.state.isSelectScrollXThumb = false;
    this.state.isSelectScrollYThumb = false;
    this.state.isSelectRulerX = false;
    this.state.isSelectRulerY = false;
    this.state.isResizeColumn = false;
    this.state.isResizeRow = false;
    this.resizeColumnCell = null;
    this.resizeRowCell = null;
    this.drawScrollBar(this.ctx);
    // this.mousePoint = null;
    if (
      this.state.unSelectCell ||
      this.state.unSelectRulerX ||
      this.state.unSelectRulerY
    ) {
      this.reCalcActiveArr();
      this.unActiveRange = null;
      this.state.unSelectCell = false;
      this.state.unSelectRulerX = false;
      this.state.unSelectRulerY = false;
      this.setActive();
      this.drawScrollBar(this.ctx);
      this.drawRuler(this.ctx);
    }
  }

  onMouseEnter(event) {
    // console.log('mouseenter', event);
  }

  onMouseLeave(event) {
    // console.log('mouseleave', event);
  }

  onContextMenu(event: MouseEvent) {
    event.returnValue = false;
  }

  inCellArea(x: number, y: number, cell: Cell) {
    return (
      inRange(
        x,
        cell.x - this.scrollLeft,
        cell.x - this.scrollLeft + cell.width,
        true
      ) &&
      inRange(
        y,
        cell.y - this.scrollTop,
        cell.y - this.scrollTop + cell.height,
        true
      )
    );
  }

  inCellsArea(x: number, y: number) {
    return (
      inRange(x, this.offsetLeft, this.offsetLeft + this.clientWidth, true) &&
      inRange(y, this.offsetTop, this.offsetTop + this.clientHeight, true)
    );
  }

  inRulerXArea(x: number, y: number) {
    return (
      inRange(x, this.offsetLeft, this.width, true) &&
      inRange(y, 0, this.offsetTop, true)
    );
  }

  inRulerXResizeGap(x: number, y: number) {
    if (this.inRulerXArea(x, y)) {
      for (let i = 1, len = this.columns.length; i < len; i++) {
        if (
          inRange(
            x,
            this.columns[i].x -
              this.scrollLeft +
              this.columns[i].width -
              2 * Style.cellBorderWidth,
            this.columns[i].x -
              this.scrollLeft +
              this.columns[i].width +
              2 * Style.cellBorderWidth
          )
        ) {
          return true;
        }
      }
    }
    return false;
  }

  inRulerYResizeGap(x: number, y: number) {
    if (this.inRulerYArea(x, y)) {
      for (let i = 1, len = this.rows.length; i < len; i++) {
        if (
          inRange(
            y,
            this.rows[i].y -
              this.scrollTop +
              this.rows[i].height -
              2 * Style.cellBorderWidth,
            this.rows[i].y -
              this.scrollTop +
              this.rows[i].height +
              2 * Style.cellBorderWidth
          )
        ) {
          return true;
        }
      }
    }
    return false;
  }

  inRulerYArea(x: number, y: number) {
    return (
      inRange(x, 0, this.offsetLeft, true) &&
      inRange(y, this.offsetTop, this.height, true)
    );
  }

  inScrollYBarArea(x: number, y: number) {
    return (
      inRange(x, this.offsetLeft + this.clientWidth, this.width, true) &&
      inRange(y, this.offsetTop, this.offsetTop + this.clientHeight, true)
    );
  }
  inScrollXBarArea(x: number, y: number) {
    return (
      inRange(x, this.offsetLeft, this.width, true) &&
      inRange(y, this.offsetTop + this.clientHeight, this.height, true)
    );
  }

  inThumbAreaOfScrollBarX(x: number, y: number, judegedInScrollBarX = false) {
    if (!judegedInScrollBarX && !this.inScrollXBarArea(x, y)) {
      return;
    }
    const scrollXThumbHeight = this.getScrollXThumbHeight();
    const scrollXThumbLeft = this.getScrollXThumbLeft(scrollXThumbHeight);

    return inRange(x, scrollXThumbLeft, scrollXThumbLeft + scrollXThumbHeight);
  }

  getScrollXThumbHeight() {
    let scrollXThumbHeight =
      (this.clientWidth / this.scrollWidth) * this.clientWidth;
    if (scrollXThumbHeight < Style.scrollBarThumbMinSize) {
      scrollXThumbHeight = Style.scrollBarThumbMinSize;
    }
    return scrollXThumbHeight;
  }

  getScrollXThumbLeft(scrollXThumbHeight: number) {
    return (
      this.offsetLeft +
      (this.scrollLeft * (this.clientWidth - scrollXThumbHeight)) /
        (this.scrollWidth - this.clientWidth)
    );
  }

  inThumbAreaOfScrollBarY(x: number, y: number, judegedInScrollBarY = false) {
    if (!judegedInScrollBarY && !this.inScrollYBarArea(x, y)) {
      return false;
    }
    const scrollYThumbHeight = this.getScrollYThumbHeight();
    const scrollYThumbTop = this.getScrollYThumbTop(scrollYThumbHeight);

    return inRange(y, scrollYThumbTop, scrollYThumbTop + scrollYThumbHeight);
  }

  getScrollYThumbHeight() {
    let scrollYThumbHeight =
      (this.clientHeight / this.scrollHeight) * this.clientHeight;
    if (scrollYThumbHeight < Style.scrollBarThumbMinSize) {
      scrollYThumbHeight = Style.scrollBarThumbMinSize;
    }
    return scrollYThumbHeight;
  }

  getScrollYThumbTop(scrollYThumbHeight: number) {
    return (
      this.offsetTop +
      (this.scrollTop * (this.clientHeight - scrollYThumbHeight)) /
        (this.scrollHeight - this.clientHeight)
    );
  }

  inSelectAllArea(x: number, y: number) {
    return (
      inRange(x, 0, this.offsetLeft, true) &&
      inRange(y, 0, this.offsetTop, true)
    );
  }

  onMouseDown(event: MouseEvent) {
    this.panel.nativeElement.focus();
    event.preventDefault();
    console.log('mousedown', event);

    if (this.state.isCellEdit) {
      this.editCellCompelte();
    }

    if (event.button === 2) {
      event.returnValue = false;
      return;
    }
    const currentTime = new Date().getTime();
    let isDblClick = false;
    if (currentTime - this.mousePoint.lastModifyTime < 300) {
      isDblClick = true;
    }
    this.mousePoint = {
      x: event.clientX,
      y: event.clientY,
      lastModifyTime: new Date().getTime(),
    };
    if (this.inSelectAllArea(event.clientX, event.clientY)) {
      console.log('all');
      this.activeArr = [
        {
          rowStart: 1,
          rowEnd: Infinity,
          columnStart: 1,
          columnEnd: Infinity,
        },
      ];
      this.activeCellPos = { row: 1, column: 1, rangeIndex: 0 };
      this.setActive();
      this.drawRuler(this.ctx);
      // this.drawScrollBar(ctx);
    } else if (this.inRulerXArea(event.clientX, event.clientY)) {
      console.log('rulerx');

      // const canvas = document.createElement('canvas');
      // canvas.width = this.width;
      // canvas.height = this.height;
      // const ctx = canvas.getContext('2d');
      for (let i = 1, len = this.viewCells[0].length; i < len; i++) {
        // ctx.clearRect(0, 0, this.width, this.height);
        // ctx.rect(
        //   this.viewCells[0][i].x - this.scrollLeft,
        //   this.viewCells[0][i].y,
        //   this.viewCells[0][i].width,
        //   this.viewCells[0][i].height
        // );
        // if (ctx.isPointInPath(event.clientX, event.clientY)) {
        if (
          inRange(
            event.x,
            this.viewCells[0][i].x -
              this.scrollLeft +
              (i ? Style.cellBorderWidth : 0),
            this.viewCells[0][i].x -
              this.scrollLeft +
              this.viewCells[0][i].width -
              (i ? 2 : 1) * Style.cellBorderWidth
          )
        ) {
          const isUnActive =
            this.activeArr.some(
              (range) =>
                range.rowStart === 1 &&
                range.rowEnd === Infinity &&
                inRange(
                  this.viewCells[0][i].position.column,
                  range.columnStart,
                  range.columnEnd,
                  true
                )
            ) && event.ctrlKey;
          if (isUnActive) {
            this.state.unSelectRulerX = true;
            this.unActiveRange = {
              rowStart: 1,
              rowEnd: Infinity,
              columnStart: this.viewCells[0][i].position.column,
              columnEnd: this.viewCells[0][i].position.column,
            };
          } else {
            this.state.isSelectRulerX = true;
            this.activeArr.push({
              rowStart: 1,
              rowEnd: Infinity,
              columnStart: event.shiftKey
                ? this.activeArr.length &&
                  this.activeArr[this.activeArr.length - 1].columnStart
                : this.viewCells[0][i].position.column,
              columnEnd: this.viewCells[0][i].position.column,
            });
            this.activeCellPos = {
              row: this.activeArr[this.activeArr.length - 1].rowStart,
              column: this.activeArr[this.activeArr.length - 1].columnStart,
              rangeIndex: this.activeArr.length - 1,
            };
          }
          if (!event.ctrlKey || event.shiftKey) {
            this.activeArr = [this.activeArr[this.activeArr.length - 1]];
            this.activeCellPos.rangeIndex = 0;
          }
          this.setActive();
          this.drawRuler(this.ctx);
          break;
        } else if (
          inRange(
            event.x,
            this.viewCells[0][i].x -
              this.scrollLeft +
              this.viewCells[0][i].width -
              2 * Style.cellBorderWidth,
            this.viewCells[0][i].x -
              this.scrollLeft +
              this.viewCells[0][i].width +
              2 * Style.cellBorderWidth
          )
        ) {
          this.resizeColumnCell = this.viewCells[0][i];
          this.state.isResizeColumn = true;
        }
      }
    } else if (this.inRulerYArea(event.clientX, event.clientY)) {
      console.log('rulery');
      // const canvas = document.createElement('canvas');
      // canvas.width = this.width;
      // canvas.height = this.height;
      // const ctx = canvas.getContext('2d');
      const rowCells = this.viewCells.map((row) => row[0]);
      for (let i = 1, len = rowCells.length; i < len; i++) {
        // ctx.clearRect(0, 0, this.width, this.height);
        // ctx.rect(
        //   rowCells[i].x,
        //   rowCells[i].y - this.scrollTop,
        //   rowCells[i].width,
        //   rowCells[i].height
        // );
        // if (ctx.isPointInPath(event.clientX, event.clientY)) {
        if (
          inRange(
            event.clientY,
            rowCells[i].y - this.scrollTop + (i ? Style.cellBorderWidth : 0),
            rowCells[i].y -
              this.scrollTop +
              rowCells[i].height -
              (i ? 2 : 1) * Style.cellBorderWidth
          )
        ) {
          const isUnActive =
            this.activeArr.some(
              (range) =>
                range.columnStart === 1 &&
                range.columnEnd === Infinity &&
                inRange(
                  rowCells[i].position.row,
                  range.rowStart,
                  range.rowEnd,
                  true
                )
            ) && event.ctrlKey;
          if (isUnActive) {
            this.state.unSelectRulerY = true;
            this.unActiveRange = {
              rowStart: rowCells[i].position.row,
              rowEnd: rowCells[i].position.row,
              columnStart: 1,
              columnEnd: Infinity,
            };
          } else {
            console.log(true);
            this.state.isSelectRulerY = true;
            this.activeArr.push({
              rowStart: event.shiftKey
                ? this.activeArr.length &&
                  this.activeArr[this.activeArr.length - 1].rowStart
                : rowCells[i].position.row,
              rowEnd: rowCells[i].position.row,
              columnStart: 1,
              columnEnd: Infinity,
            });
            this.activeCellPos = {
              row: this.activeArr[this.activeArr.length - 1].rowStart,
              column: this.activeArr[this.activeArr.length - 1].columnStart,
              rangeIndex: this.activeArr.length - 1,
            };
          }
          if (!event.ctrlKey || event.shiftKey) {
            this.activeArr = [this.activeArr[this.activeArr.length - 1]];
            this.activeCellPos.rangeIndex = 0;
          }
          this.setActive();
          this.drawRuler(this.ctx);
          break;
        } else if (
          inRange(
            event.clientY,
            rowCells[i].y -
              this.scrollTop +
              rowCells[i].height -
              2 * Style.cellBorderWidth,
            rowCells[i].y -
              this.scrollTop +
              rowCells[i].height +
              2 * Style.cellBorderWidth
          )
        ) {
          console.log('dblclick', isDblClick, event);
          if (isDblClick) {
            this.resizeRow(
              rowCells[i].position.row,
              Style.cellHeight - rowCells[i].height
            );
            return;
          } else {
            this.resizeRowCell = rowCells[i];
            this.state.isResizeRow = true;
          }
        }
      }
    } else if (this.inScrollXBarArea(event.clientX, event.clientY)) {
      console.log('scrollx');
      if (this.inThumbAreaOfScrollBarX(event.clientX, event.clientY, true)) {
        this.state.isSelectScrollXThumb = true;
      } else if (
        inRange(
          event.clientX,
          this.offsetLeft,
          this.getScrollXThumbLeft(this.getScrollXThumbHeight())
        )
      ) {
        this.scrollX(-1 * Style.cellWidth);
        this.setActive();
        this.drawRuler(this.ctx);
      } else {
        this.scrollX(Style.cellWidth);
        this.setActive();
        this.drawRuler(this.ctx);
      }
    } else if (this.inScrollYBarArea(event.clientX, event.clientY)) {
      console.log('scrolly');
      if (this.inThumbAreaOfScrollBarY(event.clientX, event.clientY, true)) {
        this.state.isSelectScrollYThumb = true;
      } else if (
        inRange(
          event.clientY,
          this.offsetTop,
          this.getScrollYThumbTop(this.getScrollYThumbHeight())
        )
      ) {
        this.scrollY(-1 * Style.cellHeight);
        this.setActive();
        this.drawRuler(this.ctx);
      } else {
        this.scrollY(Style.cellHeight);
        this.setActive();
        this.drawRuler(this.ctx);
      }
    } else if (this.inCellsArea(event.clientX, event.clientY)) {
      // const canvas = document.createElement('canvas');
      // canvas.width = this.width;
      // canvas.height = this.height;
      // const ctx = canvas.getContext('2d');
      for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
        for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
          const cell = this.viewCells[i][j];
          // ctx.clearRect(0, 0, this.width, this.height);
          // ctx.rect(
          //   cell.x - this.scrollLeft,
          //   cell.y - this.scrollTop,
          //   cell.width,
          //   cell.height
          // );
          // if (ctx.isPointInPath(event.clientX, event.clientY)) {
          if (this.inCellArea(event.clientX, event.clientY, cell)) {
            const isUnActive =
              this.activeArr.some(
                (range) =>
                  inRange(
                    cell.position.row,
                    range.rowStart,
                    range.rowEnd,
                    true
                  ) &&
                  inRange(
                    cell.position.column,
                    range.columnStart,
                    range.columnEnd,
                    true
                  )
              ) && event.ctrlKey;
            if (isUnActive) {
              this.state.unSelectCell = true;
              this.unActiveRange = {
                rowStart: cell.position.row,
                rowEnd: cell.position.row,
                columnStart: cell.position.column,
                columnEnd: cell.position.column,
              };
            } else {
              if (
                cell.position.row === this.activeCellPos.row &&
                cell.position.column === this.activeCellPos.column &&
                this.activeArr[this.activeCellPos.rangeIndex].rowStart ===
                  this.activeArr[this.activeCellPos.rangeIndex].rowEnd &&
                this.activeArr[this.activeCellPos.rangeIndex].columnStart ===
                  this.activeArr[this.activeCellPos.rangeIndex].columnEnd
              ) {
                this.resetCellPerspective(cell);
                this.editingCell = cell;
                this.state.isCellEdit = true;
                this.activeArr = [
                  {
                    rowStart: cell.position.row,
                    rowEnd: cell.position.row,
                    columnStart: cell.position.column,
                    columnEnd: cell.position.column,
                  },
                ];
                this.activeCellPos.rangeIndex = 0;
              } else {
                this.state.isSelectCell = true;
                this.activeArr.push({
                  rowStart:
                    (event.shiftKey &&
                      this.activeArr.length &&
                      this.activeArr[this.activeArr.length - 1].rowStart) ||
                    cell.position.row,
                  columnStart:
                    (event.shiftKey &&
                      this.activeArr.length &&
                      this.activeArr[this.activeArr.length - 1].columnStart) ||
                    cell.position.column,
                  rowEnd: cell.position.row,
                  columnEnd: cell.position.column,
                });
                this.activeCellPos = {
                  row: this.activeArr[this.activeArr.length - 1].rowStart,
                  column: this.activeArr[this.activeArr.length - 1].columnStart,
                  rangeIndex: this.activeArr.length - 1,
                };
                if (!event.ctrlKey || event.shiftKey) {
                  this.activeArr = [this.activeArr[this.activeArr.length - 1]];
                  this.activeCellPos.rangeIndex = 0;
                }
              }
            }

            this.autoScroll(this.mousePoint.x, this.mousePoint.y);
            this.setActive();
            this.drawRuler(this.ctx);
            return;
          }
        }
      }
    }
  }

  // @throttle(20)
  onMouseMove(event: MouseEvent) {
    if (
      (!this.inCellsArea(event.clientX, event.clientY) &&
        !this.state.isSelectCell) ||
      this.state.isSelectScrollYThumb ||
      this.state.isSelectScrollXThumb ||
      this.state.isResizeColumn
    ) {
      if (
        this.state.isResizeColumn ||
        (!this.state.isSelectRulerX &&
          !this.state.isSelectRulerY &&
          this.inRulerXResizeGap(event.clientX, event.clientY))
      ) {
        this.panel.nativeElement.style.cursor = 'col-resize';
      } else if (
        this.state.isResizeRow ||
        (!this.state.isSelectRulerX &&
          !this.state.isSelectRulerY &&
          this.inRulerYResizeGap(event.clientX, event.clientY))
      ) {
        this.panel.nativeElement.style.cursor = 'row-resize';
      } else {
        this.panel.nativeElement.style.cursor = 'default';
      }
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
      this.drawScrollBar(this.ctx);
    }

    if (!this.isTicking) {
      requestAnimationFrame(() => {
        let range;
        if (this.state.isSelectCell) {
          range = this.calcActive(event.clientX, event.clientY);
        } else if (this.state.unSelectCell) {
          range = this.calcUnActive(event.clientX, event.clientY);
        } else if (this.state.isSelectScrollYThumb) {
          this.calcScrollY(event.clientX, event.clientY);
        } else if (this.state.isSelectScrollXThumb) {
          this.calcScrollX(event.clientX, event.clientY);
        } else if (this.state.isSelectRulerX) {
          range = this.calcAcitiveRulerX(event.clientX, event.clientY);
        } else if (this.state.isSelectRulerY) {
          range = this.calcActiveRulerY(event.clientX, event.clientY);
        } else if (this.state.unSelectRulerX) {
          range = this.calcUnActiveRulerX(event.clientX, event.clientY);
        } else if (this.state.unSelectRulerY) {
          range = this.calcUnActiveRulerY(event.clientX, event.clientY);
        } else if (this.state.isResizeColumn) {
          this.resizeColumn(
            this.resizeColumnCell.position.column,
            event.clientX - this.mousePoint.x
          );
        } else if (this.state.isResizeRow) {
          this.resizeRow(
            this.resizeRowCell.position.row,
            event.clientY - this.mousePoint.y
          );
        }
        if (range) {
          if (
            !this.state.unSelectCell &&
            !this.state.unSelectRulerX &&
            !this.state.unSelectRulerY
          ) {
            this.activeArr.push(range);
            this.activeArr.splice(this.activeArr.length - 2, 1);
          } else if (range) {
            this.unActiveRange = range;
          }
          this.setActive();
          this.drawRuler(this.ctx);
        }

        this.mousePoint = { x: event.clientX, y: event.clientY };
        this.isTicking = false;
      });
    }

    this.isTicking = true;
  }

  resizeColumn(column: number, deltaWidth: number) {
    this.columns[column].width += deltaWidth;
    this.columns.forEach((col, index) => {
      if (index > column) {
        col.x += deltaWidth;
      }
    });
    this.refreshView();
  }

  resizeRow(row: number, deltaHeight: number) {
    this.rows[row].height += deltaHeight;
    this.rows.forEach((r, i) => {
      if (i > row) {
        r.y += deltaHeight;
      }
    });
    this.refreshView();
  }

  autoScroll(x: number, y: number) {
    if (
      x === this.mousePoint.x &&
      y === this.mousePoint.y &&
      this.state.isSelectCell
    ) {
      if (y > this.offsetTop + this.clientHeight || y < this.offsetTop) {
        this.scrollY(
          y < this.offsetTop ? -1 * Style.cellHeight : Style.cellHeight
        );
        const range = this.calcActive(x, y);
        if (range) {
          this.activeArr.push(range);
          this.activeArr.splice(this.activeArr.length - 2, 1);
          this.setActive();
        }
      }
      if (x > this.offsetLeft + this.clientWidth || x < this.offsetLeft) {
        this.scrollX(
          x < this.offsetLeft ? -1 * Style.cellWidth : Style.cellWidth
        );
        const range = this.calcActive(this.mousePoint.x, this.mousePoint.y);
        if (range) {
          this.activeArr.push(range);
          this.activeArr.splice(this.activeArr.length - 2, 1);
          this.setActive();
        }
      }
      this.autoScrollTimeoutID = setTimeout(
        () =>
          this.mousePoint &&
          this.state.isSelectCell &&
          this.autoScroll(this.mousePoint.x, this.mousePoint.y),
        50
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
        50
      );
    }
  }

  calcAcitiveRulerX(x: number, y: number) {
    for (let i = 1, len = this.viewCells[0].length; i < len; i++) {
      const cx = this.viewCells[0][i].x - this.scrollLeft;
      const cw = this.viewCells[0][i].width;
      if (inRange(x, cx, cx + cw, true)) {
        return {
          rowStart: 1,
          rowEnd: Infinity,
          columnStart: this.activeArr[this.activeArr.length - 1].columnStart,
          columnEnd: this.viewCells[0][i].position.column,
        };
      }
    }
    // const canvas = document.createElement('canvas');
    // canvas.width = this.width;
    // canvas.height = this.height;
    // const ctx = canvas.getContext('2d');
    // for (let i = 1, len = this.viewCells[0].length; i < len; i++) {
    //   ctx.clearRect(0, 0, this.width, this.height);
    //   ctx.rect(
    //     this.viewCells[0][i].x - this.scrollLeft,
    //     this.viewCells[0][i].y,
    //     this.viewCells[0][i].width,
    //     this.viewCells[0][i].height
    //   );
    //   if (ctx.isPointInPath(x, y)) {
    //     this.setActive({
    //       rowStart: 1,
    //       rowEnd: Infinity,
    //       columnStart: this.activeRange.columnStart,
    //       columnEnd: this.viewCells[0][i].position.column,
    //     });
    //     return;
    //   }
    // }
  }

  calcUnActiveRulerX(x: number, y: number) {
    for (let i = 1, len = this.viewCells[0].length; i < len; i++) {
      const cx = this.viewCells[0][i].x - this.scrollLeft;
      const cw = this.viewCells[0][i].width;
      if (inRange(x, cx, cx + cw, true)) {
        return {
          rowStart: 1,
          rowEnd: Infinity,
          columnStart: this.unActiveRange.columnStart,
          columnEnd: this.viewCells[0][i].position.column,
        };
      }
    }
  }

  calcActiveRulerY(x: number, y: number) {
    const rowCells = this.viewCells.map((row) => row[0]);
    for (let i = 1, len = rowCells.length; i < len; i++) {
      const cy = rowCells[i].y - this.scrollTop;
      const ch = rowCells[i].height;
      if (inRange(y, cy, cy + ch, true)) {
        return {
          rowStart: this.activeArr[this.activeArr.length - 1].rowStart,
          rowEnd: rowCells[i].position.row,
          columnStart: 1,
          columnEnd: Infinity,
        };
      }
    }
    // const canvas = document.createElement('canvas');
    // canvas.width = this.width;
    // canvas.height = this.height;
    // const ctx = canvas.getContext('2d');
    // const rowCells = this.viewCells.map((row) => row[0]);
    // for (let i = 1, len = rowCells.length; i < len; i++) {
    //   ctx.clearRect(0, 0, this.width, this.height);
    //   ctx.rect(
    //     rowCells[i].x,
    //     rowCells[i].y - this.scrollTop,
    //     rowCells[i].width,
    //     rowCells[i].height
    //   );
    //   if (ctx.isPointInPath(x, y)) {
    //     this.setActive({
    //       rowStart: this.activeRange.rowStart,
    //       rowEnd: rowCells[i].position.row,
    //       columnStart: 1,
    //       columnEnd: Infinity,
    //     });
    //     return;
    //   }
    // }
  }

  calcUnActiveRulerY(x: number, y: number) {
    const rowCells = this.viewCells.map((row) => row[0]);
    for (let i = 1, len = rowCells.length; i < len; i++) {
      const cy = rowCells[i].y - this.scrollTop;
      const ch = rowCells[i].height;
      if (inRange(y, cy, cy + ch, true)) {
        return {
          rowStart: this.unActiveRange.rowStart,
          rowEnd: rowCells[i].position.row,
          columnStart: 1,
          columnEnd: Infinity,
        };
      }
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
    // const canvas = document.createElement('canvas');
    // canvas.width = this.width;
    // canvas.height = this.height;
    // const ctx = canvas.getContext('2d');
    for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
      for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
        const cell = this.viewCells[i][j];
        // ctx.clearRect(0, 0, this.width, this.height);
        // ctx.rect(
        //   cell.x - this.scrollLeft,
        //   cell.y - this.scrollTop,
        //   cell.width,
        //   cell.height
        // );
        // if (ctx.isPointInPath(x, y)) {
        if (this.inCellArea(x, y, cell)) {
          return {
            rowStart: this.activeArr[this.activeArr.length - 1].rowStart,
            columnStart: this.activeArr[this.activeArr.length - 1].columnStart,
            rowEnd: cell.position.row,
            columnEnd: cell.position.column,
          };
        }
      }
    }
  }

  calcUnActive(x: number, y: number) {
    for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
      for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
        const cell = this.viewCells[i][j];
        if (this.inCellArea(x, y, cell)) {
          return {
            rowStart: this.unActiveRange.rowStart,
            columnStart: this.unActiveRange.columnStart,
            rowEnd: cell.position.row,
            columnEnd: cell.position.column,
          };
        }
      }
    }
  }

  reCalcActiveArr() {
    const [rowStart, rowEnd, columnStart, columnEnd] = [
      Math.min(this.unActiveRange.rowStart, this.unActiveRange.rowEnd),
      Math.max(this.unActiveRange.rowStart, this.unActiveRange.rowEnd),
      Math.min(this.unActiveRange.columnStart, this.unActiveRange.columnEnd),
      Math.max(this.unActiveRange.columnStart, this.unActiveRange.columnEnd),
    ];
    const arr = this.activeArr.map((activeRange) => {
      const temp: CellRange[] = [];
      if (
        (inRange(rowStart, activeRange.rowStart, activeRange.rowEnd, true) ||
          inRange(rowEnd, activeRange.rowStart, activeRange.rowEnd, true) ||
          (rowStart <= Math.min(activeRange.rowStart, activeRange.rowEnd) &&
            rowEnd >= Math.max(activeRange.rowStart, activeRange.rowEnd))) &&
        (inRange(
          columnStart,
          activeRange.columnStart,
          activeRange.columnEnd,
          true
        ) ||
          inRange(
            columnEnd,
            activeRange.columnStart,
            activeRange.columnEnd,
            true
          ) ||
          (columnStart <=
            Math.min(activeRange.columnStart, activeRange.columnEnd) &&
            columnEnd >=
              Math.max(activeRange.columnStart, activeRange.columnEnd)))
      ) {
        if (
          inRange(
            columnStart,
            activeRange.columnStart,
            activeRange.columnEnd
          ) ||
          (columnStart >
            Math.min(activeRange.columnStart, activeRange.columnEnd) &&
            columnStart ===
              Math.max(activeRange.columnStart, activeRange.columnEnd))
        ) {
          temp.push({
            rowStart: inRange(
              rowStart,
              activeRange.rowStart,
              activeRange.rowEnd,
              true
            )
              ? rowStart
              : Math.min(activeRange.rowStart, activeRange.rowEnd),
            rowEnd: inRange(
              rowEnd,
              activeRange.rowStart,
              activeRange.rowEnd,
              true
            )
              ? rowEnd
              : Math.max(activeRange.rowStart, activeRange.rowEnd),
            columnStart: Math.min(
              activeRange.columnStart,
              activeRange.columnEnd
            ),
            columnEnd: columnStart - 1,
          });
        }

        if (
          inRange(columnEnd, activeRange.columnStart, activeRange.columnEnd) ||
          (columnEnd <
            Math.max(activeRange.columnStart, activeRange.columnEnd) &&
            columnEnd ===
              Math.min(activeRange.columnStart, activeRange.columnEnd))
        ) {
          temp.push({
            rowStart: inRange(
              rowStart,
              activeRange.rowStart,
              activeRange.rowEnd,
              true
            )
              ? rowStart
              : Math.min(activeRange.rowStart, activeRange.rowEnd),
            rowEnd: inRange(
              rowEnd,
              activeRange.rowStart,
              activeRange.rowEnd,
              true
            )
              ? rowEnd
              : Math.max(activeRange.rowStart, activeRange.rowEnd),
            columnStart: Math.min(
              columnEnd + 1,
              Math.max(activeRange.columnStart, activeRange.columnEnd)
            ),
            columnEnd: Math.max(activeRange.columnStart, activeRange.columnEnd),
          });
        }
        if (
          inRange(rowStart, activeRange.rowStart, activeRange.rowEnd) ||
          (rowStart > Math.min(activeRange.rowStart, activeRange.rowEnd) &&
            rowStart === Math.max(activeRange.rowStart, activeRange.rowEnd))
        ) {
          temp.push({
            rowStart:
              activeRange.rowStart <= activeRange.rowEnd
                ? activeRange.rowStart
                : activeRange.rowEnd,
            rowEnd: rowStart - 1,
            columnStart: activeRange.columnStart,
            columnEnd: activeRange.columnEnd,
          });
        }
        if (
          inRange(rowEnd, activeRange.rowStart, activeRange.rowEnd) ||
          (rowEnd < Math.max(activeRange.rowStart, activeRange.rowEnd) &&
            rowEnd === Math.min(activeRange.rowStart, activeRange.rowEnd))
        ) {
          temp.push({
            rowStart: rowEnd + 1,
            rowEnd:
              activeRange.rowStart <= activeRange.rowEnd
                ? activeRange.rowEnd
                : activeRange.rowStart,
            columnStart: activeRange.columnStart,
            columnEnd: activeRange.columnEnd,
          });
        }

        if (
          this.activeCellPos &&
          inRange(this.activeCellPos.row, rowStart, rowEnd, true) &&
          inRange(this.activeCellPos.column, columnStart, columnEnd, true)
        ) {
          this.activeCellPos = null;
        }
        return temp;
      } else {
        return [activeRange];
      }
    });

    this.activeArr = arr.reduce((acc, val) => acc.concat(val), []);
    if (!this.activeArr.length) {
      this.activeArr = [
        {
          rowStart: this.unActiveRange.rowStart,
          rowEnd: this.unActiveRange.rowStart,
          columnStart: this.unActiveRange.columnStart,
          columnEnd: this.unActiveRange.columnStart,
        },
      ];
    }
    this.activeArr.sort((a, b) => {
      const r = Math.min(a.rowStart, a.rowEnd) - Math.min(b.rowStart, b.rowEnd);
      if (r) {
        return r;
      } else {
        return (
          Math.min(a.columnStart, a.columnEnd) -
          Math.min(b.columnStart, b.columnEnd)
        );
      }
    });
    // if (this.activeCellPos) {
    //   this.activeArr.forEach((range, index) => {
    //     const [rStart, rEnd, cStart, cEnd] = [
    //       range.rowStart,
    //       range.rowEnd,
    //       range.columnStart,
    //       range.columnEnd,
    //     ];
    //     if (
    //       (this.activeCellPos.row === rStart ||
    //         this.activeCellPos.row === rEnd) &&
    //       (this.activeCellPos.column === cStart ||
    //         this.activeCellPos.column === cEnd)
    //     ) {
    //       if (this.activeCellPos.row !== rStart) {
    //         range.rowStart = this.activeCellPos.row;
    //         range.rowEnd = rStart;
    //       }
    //       if (this.activeCellPos.column !== cStart) {
    //         range.columnStart = this.activeCellPos.column;
    //         range.columnEnd = cStart;
    //       }
    //       this.activeCellPos.rangeIndex = index;
    //     }
    //   });
    // } else {
    let [row, column] = [Infinity, Infinity];
    this.activeArr.forEach((range) => {
      const rStart = Math.min(range.rowStart, range.rowEnd);
      const cStart = Math.min(range.columnStart, range.columnEnd);
      if (row > rStart) {
        row = rStart;
        column = cStart;
      } else if (row === rStart) {
        column = cStart < column ? cStart : column;
      }
    });
    this.activeCellPos = {
      row,
      column,
      rangeIndex: this.activeArr.findIndex(
        (range) =>
          inRange(row, range.rowStart, range.rowEnd, true) &&
          inRange(column, range.columnStart, range.columnEnd, true)
      ),
    };
    // }
  }

  onMouseUp(event: MouseEvent) {
    // console.log('onmouseup');
    this.state.isSelectCell = false;
    this.state.isSelectScrollYThumb = false;
    this.state.isSelectScrollXThumb = false;
    this.state.isSelectRulerX = false;
    this.state.isSelectRulerY = false;
    this.state.isResizeColumn = false;
    this.state.isResizeRow = false;
    this.resizeColumnCell = null;
    this.resizeRowCell = null;
    // this.mousePoint = null;
    if (
      this.state.unSelectCell ||
      this.state.unSelectRulerX ||
      this.state.unSelectRulerY
    ) {
      this.reCalcActiveArr();
      this.unActiveRange = null;
      this.state.unSelectCell = false;
      this.state.unSelectRulerX = false;
      this.state.unSelectRulerY = false;
      this.setActive();
      this.drawScrollBar(this.ctx);
      this.drawRuler(this.ctx);
    }
  }

  scrollX(deltaX: number, immediate = true) {
    if (this.scrollLeft === 0 && deltaX < 0) {
      return;
    }
    this.scrollLeft += deltaX;
    if (this.scrollLeft >= this.scrollWidth - this.clientWidth) {
      for (
        let i = 0,
          len =
            Math.ceil(
              (this.scrollLeft - this.scrollWidth + this.clientWidth) /
                Style.cellWidth
            ) + 1;
        i < len;
        i++
      ) {
        const columnLength = this.columns.length;
        this.columns.push({
          x:
            this.columns[columnLength - 1].x +
            this.columns[columnLength - 1].width,
          width: Style.cellWidth,
        });
        this.cells.forEach((row, rk) =>
          row.push(this.createCell(rk, columnLength))
        );
      }
      this.scrollWidth =
        this.cells[0][this.cells[0].length - 1].x +
        this.cells[0][this.cells[0].length - 1].width -
        this.offsetLeft;
      // this.scrollLeft = this.scrollWidth - this.clientWidth - Style.cellWidth;
    } else if (this.scrollLeft <= 0) {
      this.scrollLeft = 0;
    }

    if (immediate) {
      this.refreshView();
    }
  }

  scrollY(deltaY: number, immediate = true) {
    if (this.scrollTop === 0 && deltaY < 0) {
      return;
    }

    this.scrollTop += deltaY;
    if (this.scrollTop >= this.scrollHeight - this.clientHeight) {
      for (
        let i = 0,
          len =
            Math.ceil(
              (this.scrollTop - this.scrollHeight + this.clientHeight) /
                Style.cellHeight
            ) + 1;
        i < len;
        i++
      ) {
        this.rows.push({
          y:
            this.rows[this.rows.length - 1].y +
            this.rows[this.rows.length - 1].height,
          height: Style.cellHeight,
        });
        this.cells.push(
          Array.from({ length: this.cells[0].length }).map((cv, ck) =>
            this.createCell(this.cells.length, ck)
          )
        );
      }
      this.scrollHeight =
        this.cells[this.cells.length - 1][0].y +
        this.cells[this.cells.length - 1][0].height -
        this.offsetTop;
      // this.scrollTop = this.scrollHeight - this.clientHeight - Style.cellHeight;
    } else if (this.scrollTop <= 0) {
      this.scrollTop = 0;
    }

    if (immediate) {
      this.refreshView();
    }
  }

  onWheel(event: WheelEvent) {
    // console.log('wheel', event);
    if (!this.isTicking) {
      requestAnimationFrame(() => {
        this.scrollY(event.deltaY);

        this.isTicking = false;
      });
    }
    this.isTicking = true;
  }

  onKeyArrowUpOrDown(event: KeyboardEvent) {
    if (!event.shiftKey) {
      if (event.code === KeyCode.ArrowDown) {
        this.activeCellPos = {
          row:
            this.activeCellPos.row + 1 > this.cells.length - 1 || event.ctrlKey
              ? this.cells.length - 1
              : this.activeCellPos.row + 1,
          column: this.activeCellPos.column,
          rangeIndex: 0,
        };
      } else if (event.code === KeyCode.ArrowUp) {
        this.activeCellPos = {
          row:
            this.activeCellPos.row - 1 < 1 || event.ctrlKey
              ? 1
              : this.activeCellPos.row - 1,
          column: this.activeCellPos.column,
          rangeIndex: 0,
        };
      }
      this.activeArr = [
        {
          rowStart: this.activeCellPos.row,
          rowEnd: this.activeCellPos.row,
          columnStart: this.activeCellPos.column,
          columnEnd: this.activeCellPos.column,
        },
      ];
      if (!this.isTicking) {
        requestAnimationFrame(() => {
          this.resetCellPerspective(
            this.cells[this.activeCellPos.row][this.activeCellPos.column]
          );

          this.isTicking = false;
        });
      }
    } else {
      // let range;
      // for (let i = this.activeArr.length - 1; i >= 0; i--) {
      //   if (
      //     this.activeArr[i].rowStart === this.activeCellPos.row &&
      //     this.activeArr[i].columnStart === this.activeCellPos.column
      //   ) {
      //     range = this.activeArr[i];
      //     break;
      //   }
      // }
      const range = this.activeArr[this.activeCellPos.rangeIndex];
      if (range && range.rowEnd !== Infinity) {
        if (event.code === KeyCode.ArrowDown) {
          range.rowEnd =
            range.rowEnd + 1 >
              this.cells[this.cells.length - 1][0].position.row || event.ctrlKey
              ? this.cells[this.cells.length - 1][0].position.row
              : range.rowEnd + 1;
        } else if (event.code === KeyCode.ArrowUp) {
          range.rowEnd =
            range.rowEnd - 1 < 1 || event.ctrlKey ? 1 : range.rowEnd - 1;
        }
        if (!this.isTicking) {
          requestAnimationFrame(() => {
            this.resetCellPerspective(
              this.cells[range.rowEnd][
                range.columnEnd === Infinity ? 0 : range.columnEnd
              ]
            );

            this.isTicking = false;
          });
        }
        this.isTicking = true;
      }
    }
  }

  onKeyArrowLeftOrRight(event: KeyboardEvent) {
    if (!event.shiftKey) {
      if (event.code === KeyCode.ArrowRight) {
        this.activeCellPos = {
          row: this.activeCellPos.row,
          column:
            this.activeCellPos.column + 1 > this.cells[0].length - 1 ||
            event.ctrlKey
              ? this.cells[0].length - 1
              : this.activeCellPos.column + 1,
          rangeIndex: 0,
        };
      } else if (event.code === KeyCode.ArrowLeft) {
        this.activeCellPos = {
          row: this.activeCellPos.row,
          column:
            this.activeCellPos.column - 1 < 1 || event.ctrlKey
              ? 1
              : this.activeCellPos.column - 1,
          rangeIndex: 0,
        };
      }
      this.activeArr = [
        {
          rowStart: this.activeCellPos.row,
          rowEnd: this.activeCellPos.row,
          columnStart: this.activeCellPos.column,
          columnEnd: this.activeCellPos.column,
        },
      ];
      if (!this.isTicking) {
        requestAnimationFrame(() => {
          this.resetCellPerspective(
            this.cells[this.activeCellPos.row][this.activeCellPos.column]
          );
          this.isTicking = false;
        });
      }
    } else {
      // let range;
      // for (let i = this.activeArr.length - 1; i >= 0; i--) {
      //   if (
      //     this.activeArr[i].rowStart === this.activeCellPos.row &&
      //     this.activeArr[i].columnStart === this.activeCellPos.column
      //   ) {
      //     range = this.activeArr[i];
      //     break;
      //   }
      // }
      const range = this.activeArr[this.activeCellPos.rangeIndex];
      if (range && range.columnEnd !== Infinity) {
        if (event.code === KeyCode.ArrowRight) {
          range.columnEnd =
            range.columnEnd + 1 > this.cells[0].length - 1 || event.ctrlKey
              ? this.cells[0].length - 1
              : range.columnEnd + 1;
        } else if (event.code === KeyCode.ArrowLeft) {
          range.columnEnd =
            range.columnEnd - 1 < 1 || event.ctrlKey ? 1 : range.columnEnd - 1;
        }
        if (!this.isTicking) {
          requestAnimationFrame(() => {
            this.resetCellPerspective(
              this.cells[range.rowEnd === Infinity ? 0 : range.rowEnd][
                event.code === KeyCode.ArrowRight
                  ? Math.max(range.columnStart, range.columnEnd)
                  : Math.min(range.columnStart, range.columnEnd)
              ]
            );
            this.isTicking = false;
          });
        }
        this.isTicking = true;
      }
    }
  }

  onKeyTabOrEnter(event: KeyboardEvent) {
    if (
      this.activeArr.length === 1 &&
      this.activeArr[0].rowStart === this.activeArr[0].rowEnd &&
      this.activeArr[0].columnStart === this.activeArr[0].columnEnd
    ) {
      if (!event.shiftKey) {
        if (event.code === KeyCode.Tab) {
          this.activeCellPos = {
            row: this.activeCellPos.row,
            column:
              this.activeCellPos.column + 1 > this.cells[0].length - 1 ||
              event.ctrlKey
                ? this.cells[0].length - 1
                : this.activeCellPos.column + 1,
            rangeIndex: 0,
          };
        } else if (event.code === KeyCode.Enter) {
          this.activeCellPos = {
            row:
              this.activeCellPos.row + 1 > this.cells.length - 1 ||
              event.ctrlKey
                ? this.cells.length - 1
                : this.activeCellPos.row + 1,
            column: this.activeCellPos.column,
            rangeIndex: 0,
          };
        }
      } else {
        if (event.code === KeyCode.Tab) {
          this.activeCellPos = {
            row: this.activeCellPos.row,
            column:
              this.activeCellPos.column - 1 < 1 || event.ctrlKey
                ? 1
                : this.activeCellPos.column - 1,
            rangeIndex: 0,
          };
        } else if (event.code === KeyCode.Enter) {
          this.activeCellPos = {
            row:
              this.activeCellPos.row - 1 < 1 || event.ctrlKey
                ? 1
                : this.activeCellPos.row - 1,
            column: this.activeCellPos.column,
            rangeIndex: 0,
          };
        }
      }
      this.activeArr = [
        {
          rowStart: this.activeCellPos.row,
          rowEnd: this.activeCellPos.row,
          columnStart: this.activeCellPos.column,
          columnEnd: this.activeCellPos.column,
        },
      ];
      if (!this.isTicking) {
        requestAnimationFrame(() => {
          this.resetCellPerspective(
            this.cells[this.activeCellPos.row][this.activeCellPos.column]
          );
          this.isTicking = false;
        });
      }
    } else {
      let range = this.activeArr[this.activeCellPos.rangeIndex];
      if (event.code === KeyCode.Tab) {
        if (
          event.shiftKey
            ? this.activeCellPos.column - 1 <
              Math.min(range.columnStart, range.columnEnd)
            : this.activeCellPos.column + 1 >
              Math.max(range.columnStart, range.columnEnd)
        ) {
          if (
            event.shiftKey
              ? this.activeCellPos.row - 1 <
                Math.min(range.rowStart, range.rowEnd)
              : this.activeCellPos.row + 1 >
                Math.max(range.rowStart, range.rowEnd)
          ) {
            const rangeIndex = event.shiftKey
              ? this.activeCellPos.rangeIndex - 1 < 0
                ? this.activeArr.length - 1
                : this.activeCellPos.rangeIndex - 1
              : this.activeCellPos.rangeIndex + 1 > this.activeArr.length - 1
              ? 0
              : this.activeCellPos.rangeIndex + 1;
            range = this.activeArr[rangeIndex];
            this.activeCellPos = {
              row: event.shiftKey
                ? Math.max(range.rowStart, range.rowEnd)
                : Math.min(range.rowStart, range.rowEnd),
              column: event.shiftKey
                ? Math.max(range.columnStart, range.columnEnd)
                : Math.min(range.columnStart, range.columnEnd),
              rangeIndex,
            };
          } else {
            event.shiftKey
              ? this.activeCellPos.row--
              : this.activeCellPos.row++;
            this.activeCellPos.column = event.shiftKey
              ? Math.max(range.columnStart, range.columnEnd)
              : (this.activeCellPos.column = Math.min(
                  range.columnStart,
                  range.columnEnd
                ));
          }
        } else {
          event.shiftKey
            ? this.activeCellPos.column--
            : this.activeCellPos.column++;
        }
      } else if (event.code === KeyCode.Enter) {
        if (
          event.shiftKey
            ? this.activeCellPos.row - 1 <
              Math.min(range.rowStart, range.rowEnd)
            : this.activeCellPos.row + 1 >
              Math.max(range.rowStart, range.rowEnd)
        ) {
          if (
            event.shiftKey
              ? this.activeCellPos.column - 1 <
                Math.min(range.columnStart, range.columnEnd)
              : this.activeCellPos.column + 1 >
                Math.max(range.columnStart, range.columnEnd)
          ) {
            const rangeIndex = event.shiftKey
              ? this.activeCellPos.rangeIndex - 1 < 0
                ? this.activeArr.length - 1
                : this.activeCellPos.rangeIndex - 1
              : this.activeCellPos.rangeIndex + 1 > this.activeArr.length - 1
              ? 0
              : this.activeCellPos.rangeIndex + 1;
            range = this.activeArr[rangeIndex];
            this.activeCellPos = {
              row: event.shiftKey
                ? Math.max(range.rowStart, range.rowEnd)
                : Math.min(range.rowStart, range.rowEnd),
              column: event.shiftKey
                ? Math.max(range.columnStart, range.columnEnd)
                : Math.min(range.columnStart, range.columnEnd),
              rangeIndex,
            };
          } else {
            event.shiftKey
              ? this.activeCellPos.column--
              : this.activeCellPos.column++;
            this.activeCellPos.row = event.shiftKey
              ? Math.max(range.rowStart, range.rowEnd)
              : Math.min(range.rowStart, range.rowEnd);
          }
        } else {
          event.shiftKey ? this.activeCellPos.row-- : this.activeCellPos.row++;
        }
      }

      if (!this.isTicking) {
        requestAnimationFrame(() => {
          this.resetCellPerspective(
            this.cells[this.activeCellPos.row][this.activeCellPos.column]
          );
          this.isTicking = false;
        });
      }
    }
  }

  resetCellPerspective(cell: Cell) {
    if (cell.x - this.scrollLeft < this.offsetLeft) {
      this.scrollX(cell.x - this.scrollLeft - this.offsetLeft, false);
    } else if (
      cell.x + cell.width - this.scrollLeft >
      this.clientWidth + this.offsetLeft
    ) {
      this.scrollX(
        cell.x -
          this.scrollLeft -
          this.clientWidth -
          this.offsetLeft +
          cell.width,
        false
      );
    }
    if (cell.y - this.scrollTop < this.offsetTop) {
      this.scrollY(cell.y - this.scrollTop - this.offsetTop, false);
    } else if (
      cell.y + cell.height - this.scrollTop >
      this.clientHeight + this.offsetTop
    ) {
      this.scrollY(
        cell.y -
          this.scrollTop -
          this.clientHeight -
          this.offsetTop +
          cell.height,
        false
      );
    }
    this.refreshView();
  }

  onKeyHome(event: KeyboardEvent) {
    this.activeArr = [
      {
        rowStart:
          event.shiftKey || (!event.shiftKey && !event.ctrlKey)
            ? this.activeCellPos.row
            : 1,
        rowEnd:
          (event.shiftKey && !event.ctrlKey) ||
          (!event.shiftKey && !event.ctrlKey)
            ? this.activeCellPos.row
            : 1,
        columnStart: event.shiftKey ? this.activeCellPos.column : 1,
        columnEnd:
          event.shiftKey || event.ctrlKey || (!event.shiftKey && !event.ctrlKey)
            ? 1
            : this.activeCellPos.column,
      },
    ];
    this.activeCellPos = {
      row: this.activeArr[0].rowStart,
      column: this.activeArr[0].columnStart,
      rangeIndex: 0,
    };
    const range = this.activeArr[0];
    if (this.scrollLeft > 0) {
      this.scrollX(-1 * this.scrollLeft, false);
    }
    if (event.ctrlKey && this.scrollTop > 0) {
      this.scrollY(-1 * this.scrollTop, false);
    }
    this.refreshView();
  }

  onKeyPageUpOrDown(event: KeyboardEvent) {
    let posY =
      this.cells[this.activeCellPos.row][this.activeCellPos.column].y +
      (event.code === KeyCode.PageDown ? 1 : -1) * this.clientHeight;
    posY = posY < this.offsetTop ? this.offsetTop : posY;
    let i = this.activeCellPos.row;
    while (true) {
      const [y, height] = this.cells[i]
        ? [
            this.cells[i][this.activeCellPos.column].y,
            this.cells[i][this.activeCellPos.column].height,
          ]
        : [
            this.cells[this.cells.length - 1][0].y +
              this.cells[this.cells.length - 1][0].height +
              (i - this.cells.length) * Style.cellHeight,
            Style.cellHeight,
          ];
      console.log(i, posY, y, height);
      if (inRange(posY, y, y + height, true)) {
        this.scrollY(y - this.cells[this.activeCellPos.row][0].y, false);
        this.activeArr = [
          {
            rowStart: event.shiftKey ? this.activeCellPos.row : i,
            rowEnd: i,
            columnStart: this.activeCellPos.column,
            columnEnd: event.shiftKey
              ? this.activeArr[this.activeCellPos.rangeIndex].columnStart ===
                this.activeCellPos.column
                ? this.activeArr[this.activeCellPos.rangeIndex].columnEnd
                : this.activeArr[this.activeCellPos.rangeIndex].columnStart
              : this.activeCellPos.column,
          },
        ];
        this.activeCellPos = {
          row: event.shiftKey ? this.activeCellPos.row : i,
          column: this.activeCellPos.column,
          rangeIndex: 0,
        };
        if (!this.isTicking) {
          requestAnimationFrame(() => {
            this.resetCellPerspective(
              event.shiftKey
                ? this.cells[
                    event.code === KeyCode.PageDown
                      ? Math.max(
                          this.activeArr[0].rowStart,
                          this.activeArr[0].rowEnd
                        )
                      : Math.min(
                          this.activeArr[0].rowStart,
                          this.activeArr[0].rowEnd
                        )
                  ][this.activeArr[0].columnEnd]
                : this.cells[this.activeCellPos.row][this.activeCellPos.column]
            );
            this.isTicking = false;
          });
        }
        this.isTicking = true;
        break;
      } else {
        event.code === KeyCode.PageDown ? i++ : i--;
      }
    }
  }

  deleteContent(rangeArr: CellRange[]) {
    for (let i = rangeArr.length - 1; i >= 0; i--) {
      const range = this.activeArr[i];
      for (
        let rs = Math.min(range.rowStart, range.rowEnd),
          re = Math.min(
            Math.max(range.rowStart, range.rowEnd),
            this.cells.length - 1
          );
        rs <= re;
        rs++
      ) {
        for (
          let cs = Math.min(range.columnStart, range.columnEnd),
            ce = Math.min(
              Math.max(range.columnStart, range.columnEnd),
              this.cells[0].length - 1
            );
          cs <= ce;
          cs++
        ) {
          if (this.cells[rs][cs].content && this.cells[rs][cs].content.value) {
            this.cells[rs][cs].content.value = null;
            this.drawCell(this.ctx, this.cells[rs][cs]);
          }
        }
      }
    }
  }

  onKeyDown(event: KeyboardEvent) {
    console.log('keydown', event);
    switch (event.code) {
      case KeyCode.Tab:
      case KeyCode.Enter:
        event.preventDefault();
        this.onKeyTabOrEnter(event);
        this.cells[0][1].width = 200;
        this.refreshView();
        break;
      case KeyCode.ArrowUp:
      case KeyCode.ArrowDown:
        event.preventDefault();
        this.onKeyArrowUpOrDown(event);
        break;
      case KeyCode.ArrowLeft:
      case KeyCode.ArrowRight:
        event.preventDefault();
        this.onKeyArrowLeftOrRight(event);
        break;
      case KeyCode.Home:
        event.preventDefault();
        this.onKeyHome(event);
        break;
      case KeyCode.PageUp:
      case KeyCode.PageDown:
        event.preventDefault();
        this.onKeyPageUpOrDown(event);
        break;
      case KeyCode.Delete:
        this.deleteContent(this.activeArr);
        break;
      case KeyCode.Backspace:
        this.resetCellPerspective(
          this.cells[this.activeCellPos.row][this.activeCellPos.column]
        );
        this.editingCell = this.cells[this.activeCellPos.row][
          this.activeCellPos.column
        ];
        this.editingCell.content.value = null;
        this.state.isCellEdit = true;
        break;
      default:
        break;
    }
  }

  onKeyPress(event: KeyboardEvent) {
    console.log('keypress', event);
    event.preventDefault();
    this.resetCellPerspective(
      this.cells[this.activeCellPos.row][this.activeCellPos.column]
    );
    this.editingCell = this.cells[this.activeCellPos.row][
      this.activeCellPos.column
    ];
    this.editingCell.content.value = event.key;
    this.state.isCellEdit = true;
  }

  editCellCompelte(change = true) {
    this.state.isCellEdit = false;
    if (!change) {
      this.editingCell.content.value = this.editingCell.content.previousValue;
    } else {
      this.editingCell.content.previousValue = this.editingCell.content.value;
    }
    this.drawCell(this.ctx, this.editingCell);
    this.editingCell = null;
    this.panel.nativeElement.focus();
  }

  onEditCellKeyDown(event: KeyboardEvent) {
    switch (event.code) {
      case KeyCode.Tab:
      case KeyCode.Enter:
        this.editCellCompelte();
        event.preventDefault();
        this.onKeyTabOrEnter(event);
        break;
      case KeyCode.ArrowUp:
      case KeyCode.ArrowDown:
        if (!this.editingCell.content || !this.editingCell.content.value) {
          this.editCellCompelte();
          event.preventDefault();
          this.onKeyArrowUpOrDown(event);
        }
        break;
      case KeyCode.ArrowLeft:
      case KeyCode.ArrowRight:
        if (!this.editingCell.content || !this.editingCell.content.value) {
          this.editCellCompelte();
          event.preventDefault();
          this.onKeyArrowLeftOrRight(event);
        }
        break;
      case KeyCode.Escape:
        this.editCellCompelte(false);
        break;
      default:
        break;
    }
  }
}
