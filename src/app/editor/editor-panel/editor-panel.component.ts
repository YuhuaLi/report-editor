import { Style } from './style.const';
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
  @ViewChild('actionPanel') actionPanel: ElementRef;

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
  activeArr: any = [{ rowStart: 1, columnStart: 1, rowEnd: 1, columnEnd: 1 }];

  ctx: CanvasRenderingContext2D;
  actionCtx: CanvasRenderingContext2D;

  constructor(private elmentRef: ElementRef) {}

  ngOnInit(): void {
    this.width = this.elmentRef.nativeElement.offsetWidth;
    this.height = this.elmentRef.nativeElement.offsetHeight;

    this.clientHeight = this.height - this.offsetHeight - Style.scrollBarWidth;
  }

  ngAfterViewInit() {
    this.ctx = this.panel.nativeElement.getContext('2d');
    this.actionCtx = this.actionPanel.nativeElement.getContext('2d');
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
      return Array.from({ length: this.viewColumnCount + 1 }, (cv, ck) =>
        this.createCell(rk, ck)
      ).reduce((acc, cur) => acc.push(cur) && acc, []);
    });
    this.scrollWidth =
      this.cells[0][this.cells[0].length - 1].x +
      this.cells[0][this.cells[0].length - 1].width -
      this.offsetWidth;
    this.scrollHeight =
      this.cells[this.cells.length - 1][0].y +
      this.cells[this.cells.length - 1][0].height -
      this.offsetHeight;

    this.panel.nativeElement.width = this.width;
    this.panel.nativeElement.height = this.height;
    this.actionPanel.nativeElement.width = this.width;
    this.actionPanel.nativeElement.height = this.height;

    this.panel.nativeElement.focus();

    this.refreshView();
  }

  refreshView() {
    let startRowIndex = this.cells.slice(1).findIndex((row, index) => {
      return (
        row[0].y - this.scrollTop <= this.offsetHeight &&
        row[0].y + row[0].height - this.scrollTop >= this.offsetHeight
      );
    });
    startRowIndex = startRowIndex ? startRowIndex : startRowIndex + 1;
    let startColumnIndex = this.cells[0].slice(1).findIndex((cell, index) => {
      return (
        cell.x - this.scrollLeft <= this.offsetWidth &&
        cell.x + cell.width - this.scrollLeft >= this.offsetWidth
      );
    });
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
      position: { row: rk, column: ck },
      x: ck === 0 ? 0 : this.offsetWidth + (ck - 1) * Style.cellWidth,
      y: isXRuler ? 0 : this.offsetHeight + (rk - 1) * Style.cellHeight,
      width: isYRuler ? this.offsetWidth : Style.cellWidth,
      height: isXRuler ? this.offsetHeight : Style.cellHeight,
      type:
        isXRuler && isYRuler ? 'all' : isXRuler || isYRuler ? 'ruler' : 'cell',
      content: {
        value:
          isYRuler && !isXRuler
            ? this.generateRowNum(rk)
            : isXRuler && !isYRuler
            ? this.generateColumnNum(ck)
            : this.generateRowNum(rk) + this.generateColumnNum(ck),
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
      this.offsetWidth,
      this.height - Style.scrollBarWidth,
      this.width - Style.scrollBarWidth,
      this.height
    );
    ctx.strokeRect(
      this.offsetWidth,
      this.height - Style.scrollBarWidth,
      this.width - Style.scrollBarWidth,
      this.height
    );
    ctx.fillRect(
      this.width - Style.scrollBarWidth,
      this.offsetHeight,
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
      this.offsetHeight,
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
    ctx.fillRect(0, 0, this.offsetWidth, this.offsetHeight);
    ctx.strokeRect(0, 0, this.offsetWidth, this.offsetHeight);
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
            2 * Style.cellBorderWidth,
          this.cells[rowStart][columnStart].y -
            this.scrollTop +
            2 * Style.cellBorderWidth,
          this.cells[rowEnd][columnEnd].x -
            this.cells[rowStart][columnStart].x +
            this.cells[rowEnd][columnEnd].width -
            4 * Style.cellBorderWidth,
          this.cells[rowEnd][columnEnd].y -
            this.cells[rowStart][columnStart].y +
            this.cells[rowEnd][columnEnd].height -
            4 * Style.cellBorderWidth
        );
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
    }
    this.actionCtx.drawImage(
      canvas,
      this.offsetWidth,
      this.offsetHeight,
      this.clientWidth,
      this.clientHeight,
      this.offsetWidth,
      this.offsetHeight,
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
  //       this.offsetWidth,
  //       this.offsetHeight,
  //       this.clientWidth,
  //       this.clientHeight,
  //       this.offsetWidth,
  //       this.offsetHeight,
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
    console.log('mouseover', event);
  }

  onMouseOut(event: MouseEvent) {
    this.state.isSelectCell = false;
    this.state.isScrollXThumbHover = false;
    this.state.isScrollYThumbHover = false;
    this.state.isSelectScrollXThumb = false;
    this.state.isSelectScrollYThumb = false;
    this.state.isSelectRulerX = false;
    this.state.isSelectRulerY = false;
    this.drawScrollBar(this.ctx);
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
      this.offsetWidth +
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
      this.offsetHeight +
      (this.scrollTop * (this.clientHeight - scrollYThumbHeight)) /
        (this.scrollHeight - this.clientHeight)
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
    this.mousePoint = { x: event.clientX, y: event.clientY };
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
      this.setActive();
      this.drawRuler(this.ctx);
      // this.drawScrollBar(ctx);
    } else if (this.inRulerXArea(event.clientX, event.clientY)) {
      console.log('rulerx');
      this.state.isSelectRulerX = true;
      const canvas = document.createElement('canvas');
      canvas.width = this.width;
      canvas.height = this.height;
      const ctx = canvas.getContext('2d');
      for (let i = 1, len = this.viewCells[0].length; i < len; i++) {
        ctx.clearRect(0, 0, this.width, this.height);
        ctx.rect(
          this.viewCells[0][i].x - this.scrollLeft,
          this.viewCells[0][i].y,
          this.viewCells[0][i].width,
          this.viewCells[0][i].height
        );
        if (ctx.isPointInPath(event.clientX, event.clientY)) {
          this.activeArr.push({
            rowStart: 1,
            rowEnd: Infinity,
            columnStart: event.shiftKey
              ? this.activeArr.length &&
                this.activeArr[this.activeArr.length - 1].columnStart
              : this.viewCells[0][i].position.column,
            columnEnd: this.viewCells[0][i].position.column,
          });
          if (!event.ctrlKey) {
            this.activeArr = [this.activeArr[this.activeArr.length - 1]];
          }
          this.setActive();
          this.drawRuler(this.ctx);
          break;
        }
      }
    } else if (this.inRulerYArea(event.clientX, event.clientY)) {
      this.state.isSelectRulerY = true;
      console.log('rulery');
      const canvas = document.createElement('canvas');
      canvas.width = this.width;
      canvas.height = this.height;
      const ctx = canvas.getContext('2d');
      const rowCells = this.viewCells.map((row) => row[0]);
      for (let i = 1, len = rowCells.length; i < len; i++) {
        ctx.clearRect(0, 0, this.width, this.height);
        ctx.rect(
          rowCells[i].x,
          rowCells[i].y - this.scrollTop,
          rowCells[i].width,
          rowCells[i].height
        );
        if (ctx.isPointInPath(event.clientX, event.clientY)) {
          this.activeArr.push({
            rowStart: event.shiftKey
              ? this.activeArr.length &&
                this.activeArr[this.activeArr.length - 1].rowStart
              : rowCells[i].position.row,
            rowEnd: rowCells[i].position.row,
            columnStart: 1,
            columnEnd: Infinity,
          });
          if (!event.ctrlKey) {
            this.activeArr = [this.activeArr[this.activeArr.length - 1]];
          }
          this.setActive();
          this.drawRuler(this.ctx);
          break;
        }
      }
    } else if (this.inScrollXBarArea(event.clientX, event.clientY)) {
      console.log('scrollx');
      if (this.inThumbAreaOfScrollBarX(event.clientX, event.clientY, true)) {
        this.state.isSelectScrollXThumb = true;
      } else if (
        inRange(
          event.clientX,
          this.offsetWidth,
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
          this.offsetHeight,
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
    } else if (this.inCellArea(event.clientX, event.clientY)) {
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
            console.log(this.activeArr);
            if (!event.ctrlKey) {
              this.activeArr = [this.activeArr[this.activeArr.length - 1]];
            }

            this.state.isSelectCell = true;
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
      (!this.inCellArea(event.clientX, event.clientY) &&
        !this.state.isSelectCell) ||
      this.state.isSelectScrollYThumb ||
      this.state.isSelectScrollXThumb
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
      this.drawScrollBar(this.ctx);
    }

    if (!this.isTicking) {
      requestAnimationFrame(() => {
        let range;
        if (this.state.isSelectCell) {
          range = this.calcActive(event.clientX, event.clientY);
        } else if (this.state.isSelectScrollYThumb) {
          this.calcScrollY(event.clientX, event.clientY);
        } else if (this.state.isSelectScrollXThumb) {
          this.calcScrollX(event.clientX, event.clientY);
        } else if (this.state.isSelectRulerX) {
          range = this.calcAcitiveRulerX(event.clientX, event.clientY);
        } else if (this.state.isSelectRulerY) {
          range = this.calcAcitiveRulerY(event.clientX, event.clientY);
        }

        if (range) {
          this.activeArr.push(range);
          this.activeArr.splice(this.activeArr.length - 2, 1);
          this.setActive();
          this.drawRuler(this.ctx);
        }

        this.mousePoint = { x: event.clientX, y: event.clientY };
        this.isTicking = false;
      });
    }

    this.isTicking = true;
  }

  autoScroll(x: number, y: number) {
    if (
      x === this.mousePoint.x &&
      y === this.mousePoint.y &&
      this.state.isSelectCell
    ) {
      if (y > this.offsetHeight + this.clientHeight || y < this.offsetHeight) {
        this.scrollY(
          y < this.offsetHeight ? -1 * Style.cellHeight : Style.cellHeight
        );
        const range = this.calcActive(x, y);
        if (range) {
          this.activeArr.push(range);
          this.activeArr.splice(this.activeArr.length - 2, 1);
          this.setActive();
        }
      }
      if (x > this.offsetWidth + this.clientWidth || x < this.offsetWidth) {
        this.scrollX(
          x < this.offsetWidth ? -1 * Style.cellWidth : Style.cellWidth
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

  calcAcitiveRulerY(x: number, y: number) {
    const rowCells = this.viewCells.map((row) => row[0]);
    for (let i = 1, len = rowCells.length; i < len; i++) {
      const cy = rowCells[i].y - this.scrollTop;
      const ch = rowCells[i].height;
      if (inRange(y, cy, cy + ch, true)) {
        return {
          rowStart: this.activeArr[this.activeArr.length].rowStart,
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

  onMouseUp(event: MouseEvent) {
    this.state.isSelectCell = false;
    this.state.isSelectScrollYThumb = false;
    this.state.isSelectScrollXThumb = false;
    this.state.isSelectRulerX = false;
    this.state.isSelectRulerY = false;
    this.mousePoint = null;
  }

  scrollX(deltaX: number) {
    if (this.scrollLeft === 0 && deltaX < 0) {
      return;
    }
    this.scrollLeft += deltaX;
    if (this.scrollLeft >= this.scrollWidth - this.clientWidth) {
      for (
        let i = 0, len = Math.ceil(deltaX / Style.cellWidth) + 1;
        i < len;
        i++
      ) {
        const columnLength = this.cells[0].length;
        this.cells.forEach((row, rk) =>
          row.push(this.createCell(rk, columnLength))
        );
      }
      this.scrollWidth =
        this.cells[0][this.cells[0].length - 1].x +
        this.cells[0][this.cells[0].length - 1].width -
        this.offsetWidth;
      // this.scrollLeft = this.scrollWidth - this.clientWidth - Style.cellWidth;
    } else if (this.scrollLeft <= 0) {
      this.scrollLeft = 0;
    }

    this.refreshView();
  }

  scrollY(deltaY: number) {
    if (this.scrollTop === 0 && deltaY < 0) {
      return;
    }

    this.scrollTop += deltaY;
    if (this.scrollTop >= this.scrollHeight - this.clientHeight) {
      for (
        let i = 0, len = Math.ceil(deltaY / Style.cellHeight) + 1;
        i < len;
        i++
      ) {
        this.cells.push(
          Array.from({ length: this.cells[0].length }).map((cv, ck) =>
            this.createCell(this.cells.length, ck)
          )
        );
      }
      this.scrollHeight =
        this.cells[this.cells.length - 1][0].y +
        this.cells[this.cells.length - 1][0].height -
        this.offsetHeight;
      // this.scrollTop = this.scrollHeight - this.clientHeight - Style.cellHeight;
    } else if (this.scrollTop <= 0) {
      this.scrollTop = 0;
    }

    this.refreshView();
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

  onKeyDown(event: KeyboardEvent) {
    // console.log('keydown', event);
  }
}
