import { cloneDeep } from 'lodash';
import {
  Cell,
  CellRange,
  Style,
  KeyboardCode,
  CellStyle,
} from 'src/app/core/model';
import { inRange } from 'src/app/core/utils/function';
import { FloatElement } from 'src/app/core/model/float-element';
import { element } from 'protractor';
import { throttle } from 'src/app/core/decorator/throttle.decorator';
import { LogicPosition } from 'src/app/core/model/logic-position.enmu';

export class Panel {
  multiple = 1;
  width = 0;
  height = 0;
  offsetWidth = 0;
  offsetHeight = 0;
  viewRowCount = 0;
  viewColumnCount = 0;
  offsetLeft = 60;
  offsetTop = 30;
  columns = [];
  rows = [];
  cells: Cell[][] = [];
  viewCells: Cell[][] = [];
  editingCell: Cell;
  // scrollBarWidth = this.style.scrollBarWidth;
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
  clipAnimationTimeoutID: any;
  unActiveCellPos: any;
  //   activeCellPos: any = { row: 1, column: 1, rangeIndex: 0 };
  private activeCellPos$: any;
  activeCell: Cell;
  get activeCellPos() {
    return this.activeCellPos$;
  }
  set activeCellPos(val) {
    this.activeCellPos$ = val;
    if (val) {
      this.activeCell = this.cells[val.row][val.column].isCombined
        ? this.cells[val.row][val.column].combineCell
        : this.cells[val.row][val.column];
      this.floatArr.forEach((elem) => (elem.isActive = false));
    } else {
      this.activeCell = null;
    }
  }

  activeArr: CellRange[] = [];
  unActiveRange: CellRange;
  resizeColumnCell: Cell;
  resizeRowCell: Cell;
  clipBoard: { range?: CellRange; isCut: boolean; floatElement?: FloatElement };

  floatArr: FloatElement[] = [];

  canvas: HTMLCanvasElement;
  actionCanvas: HTMLCanvasElement;
  animationCanvas: HTMLCanvasElement;
  floatCanvas: HTMLCanvasElement;
  offscreenCanvas: HTMLCanvasElement;
  floatActionCanvas: HTMLCanvasElement;

  ctx: CanvasRenderingContext2D;
  actionCtx: CanvasRenderingContext2D;
  animationCtx: CanvasRenderingContext2D;
  floatCtx: CanvasRenderingContext2D;
  offscreenCtx: CanvasRenderingContext2D;
  floatActionCtx: CanvasRenderingContext2D;

  historyState = [];
  style = new Proxy(Style, {
    get: (obj, prop: string) => {
      if (prop in obj) {
        return isNaN(obj[prop]) || !/^scrollBar/i.test(prop)
          ? obj[prop]
          : obj[prop] / this.multiple;
      }
      return null;
    },
  });

  init() {
    this.ctx = this.canvas.getContext('2d');
    this.actionCtx = this.actionCanvas.getContext('2d');
    this.animationCtx = this.animationCanvas.getContext('2d');
    this.floatCtx = this.floatCanvas.getContext('2d');
    this.floatActionCtx = this.floatActionCanvas.getContext('2d');
    this.resize();
    this.activeCellPos = { row: 1, column: 1, rangeIndex: 0 };
    this.activeArr = [{ rowStart: 1, columnStart: 1, rowEnd: 1, columnEnd: 1 }];
    this.refreshView();
    this.canvas.focus();
    // document.onpaste = this.onPaste;
    document.onmouseup = this.onMouseUp;
  }

  resize() {
    this.offsetWidth = this.canvas.offsetWidth;
    this.offsetHeight = this.canvas.offsetHeight;
    this.canvas.width = this.offsetWidth;
    this.canvas.height = this.offsetHeight;
    this.actionCanvas.width = this.offsetWidth;
    this.actionCanvas.height = this.offsetHeight;
    this.animationCanvas.width = this.offsetWidth;
    this.animationCanvas.height = this.offsetHeight;
    this.floatCanvas.width = this.offsetWidth;
    this.floatCanvas.height = this.offsetHeight;
    this.floatActionCanvas.width = this.offsetWidth;
    this.floatActionCanvas.height = this.offsetHeight;
    this.offscreenCanvas = document.createElement('canvas');
    this.offscreenCanvas.width = this.offsetWidth;
    this.offscreenCanvas.height = this.offsetHeight;
    this.width = this.offsetWidth / this.multiple;
    this.height = this.offsetHeight / this.multiple;
    this.clientWidth = this.width - this.offsetLeft - this.style.scrollBarWidth;
    this.clientHeight =
      this.height - this.offsetTop - this.style.scrollBarWidth;
    this.viewRowCount =
      Math.ceil((this.height - this.offsetTop) / this.style.cellHeight) + 2;
    if (this.viewCells.length < this.viewRowCount) {
      this.createRow(this.viewRowCount - this.viewCells.length);
    }
    this.viewColumnCount =
      Math.ceil((this.width - this.offsetLeft) / this.style.cellWidth) + 2;
    if ((this.viewCells[0] || []).length < this.viewColumnCount) {
      this.createColumn(
        this.viewColumnCount - (this.viewCells[0] || []).length
      );
    }

    this.scrollWidth =
      this.cells[0][this.cells[0].length - 1].x +
      this.cells[0][this.cells[0].length - 1].width -
      this.offsetLeft;
    this.scrollHeight =
      this.cells[this.cells.length - 1][0].y +
      this.cells[this.cells.length - 1][0].height -
      this.offsetTop;

    this.offscreenCtx = this.offscreenCanvas.getContext('2d');

    this.setTransform();
  }

  /** 重绘视图 */
  refreshView(reCalc = true) {
    let startRowIndex = this.rows
      .slice(1)
      .findIndex(
        (row, index) =>
          row.y - this.scrollTop <= this.offsetTop &&
          row.y + row.height - this.scrollTop >= this.offsetTop
      );
    startRowIndex = startRowIndex ? startRowIndex : startRowIndex + 1;
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

    if (!this.isTicking) {
      requestAnimationFrame(() => {
        this.ctx.clearRect(0, 0, this.width, this.height);

        // const canvas = document.createElement('canvas');
        // canvas.width = this.width;
        // canvas.height = this.height;
        // const ctx = canvas.getContext('2d');
        this.setActive();
        this.setClipStatusAnimation();
        this.drawFloat();
        // this.drawPanel(ctx);
        // this.drawRuler(ctx);
        // this.drawScrollBar(ctx);
        this.drawPanel(this.ctx);
        this.drawRuler(this.ctx);
        this.drawScrollBar(this.ctx);

        this.isTicking = false;
        // this.ctx.drawImage(canvas, 0, 0);
      });
    }
    this.isTicking = true;
  }

  /** 创建单元格对象 */
  createCell(rk: number, ck: number): Cell {
    const isXRuler = rk === 0;
    const isYRuler = ck === 0;
    const isCombined =
      (rk > 1 &&
        this.cells[1][ck] &&
        (this.cells[1][ck].rowSpan === Infinity ||
          (this.cells[1][ck].isCombined &&
            this.cells[1][ck].combineCell.rowSpan === Infinity))) ||
      (ck > 1 &&
        this.cells[rk] &&
        (this.cells[rk][1].colSpan === Infinity ||
          (this.cells[rk][1].isCombined &&
            this.cells[rk][1].combineCell.colSpan === Infinity)));
    return {
      columns: this.columns,
      rows: this.rows,
      position: {
        row: rk,
        column: ck,
      },
      index: this.generateColumnNum(ck) + this.generateRowNum(rk),
      get x() {
        if (
          !this.columns[this.position.column] ||
          !('x' in this.columns[this.position.column])
        ) {
          this.columns[this.position.column] = {
            x:
              this.columns[this.position.column - 1].x +
              this.columns[this.position.column - 1].width,
          };
        }
        return this.columns[this.position.column].x;
      },
      set x(val) {
        this.columns[this.position.column].x = val;
      },
      get y() {
        if (
          !this.rows[this.position.row] ||
          !('y' in this.rows[this.position.row])
        ) {
          this.rows[this.position.row] = {
            y:
              this.rows[this.position.row - 1].y +
              this.rows[this.position.row - 1].height,
          };
        }
        return this.rows[this.position.row].y;
      },
      set y(val) {
        this.rows[this.position.row].y = val;
      },
      get width() {
        let colSpan = Math.min(this.columns.length - 1, this.colSpan);
        let width = 0;
        while (colSpan-- > 0) {
          width += this.columns[this.position.column + colSpan].width;
        }
        return width;
      },
      set width(val) {
        this.columns[this.position.column].width = val;
      },
      get height() {
        let rowSpan = Math.min(this.rows.length - 1, this.rowSpan);
        let height = 0;
        while (rowSpan-- > 0) {
          height += this.rows[this.position.row + rowSpan].height;
        }
        return height;
      },
      set height(val) {
        this.rows[this.position.row].height = val;
      },
      type:
        isXRuler && isYRuler
          ? 'all'
          : (isXRuler && 'rulerX') || (isYRuler && 'rulerY') || 'cell',
      content: {
        value:
          isYRuler && !isXRuler
            ? this.generateRowNum(rk)
            : isXRuler && !isYRuler
            ? this.generateColumnNum(ck)
            : null, //this.generateColumnNum(ck) + this.generateRowNum(rk)
        previousValue: null,
        previouseHtml: null
      },
      style: {
        fontWeight:
          isXRuler || isYRuler
            ? this.style.rulerCellFontWeight
            : this.getCellDefaultAttr('fontWeight', rk, ck) ||
              this.style.cellFontWeight,
        textAlign:
          isXRuler || isYRuler
            ? (this.style.cellTextAlignCenter as CanvasTextAlign)
            : this.getCellDefaultAttr('textAlign', rk, ck) ||
              (this.style.cellTextAlignLeft as CanvasTextAlign),
        textBaseline:
          this.getCellDefaultAttr('textBaseline', rk, ck) ||
          this.style.cellTextBaseline,
        fontStyle:
          this.getCellDefaultAttr('fontStyle', rk, ck) ||
          this.style.cellFontStyle,
        fontFamily:
          isXRuler || isYRuler
            ? this.style.rulerCellFontFamily
            : this.getCellDefaultAttr('fontFamily', rk, ck) ||
              this.style.cellFontFamily,
        fontSize:
          isXRuler || isYRuler
            ? this.style.rulerCellFontSize
            : this.getCellDefaultAttr('fontSize', rk, ck) ||
              this.style.cellFontSize,
        background:
          (isXRuler && !isYRuler) || (isYRuler && !isXRuler)
            ? this.style.rulerCellBackgroundColor
            : this.getCellDefaultAttr('background', rk, ck) ||
              this.style.cellBackgroundColor,
        color:
          isXRuler || isYRuler
            ? this.style.rulerCellColor
            : this.getCellDefaultAttr('color', rk, ck) || this.style.cellColor,
        borderWidth:
          isXRuler || isYRuler
            ? this.style.rulerCellBorderWidth
            : this.getCellDefaultAttr('borderWidth', rk, ck) ||
              this.style.cellBorderWidth,
        borderColor:
          isXRuler || isYRuler
            ? this.style.rulerCellBorderColor
            : this.getCellDefaultAttr('borderColor', rk, ck) ||
              this.style.cellBorderColor,
      },
      rowSpan: 1,
      colSpan: 1,
      get isCombined() {
        return !!this.combineCell;
      },
      combineCell:
        (isCombined &&
          ((this.cells[1][ck].isCombined && this.cells[1][ck].combineCell) ||
            (this.cells[1][ck].rowSpan === Infinity && this.cells[1][ck]) ||
            (this.cells[rk] &&
              ck > 1 &&
              this.cells[rk][1].isCombined &&
              this.cells[rk][1].combineCell) ||
            (this.cells[rk][1].colSpan === Infinity && this.cells[rk][1]))) ||
        null,
      get hidden() {
        return !this.width || !this.height;
      },
    };
  }

  /** 绘制横向滚动条 */
  drawScrollBarX(ctx: CanvasRenderingContext2D) {
    ctx.save();
    ctx.fillStyle = this.style.scrollBarBackgroundColor;
    ctx.strokeStyle = this.style.scrollBarBorderColor;
    ctx.lineWidth = this.style.scrollBarBorderWidth;
    ctx.fillRect(
      this.offsetLeft,
      this.height - this.style.scrollBarWidth,
      this.width - this.style.scrollBarWidth,
      this.height
    );
    ctx.strokeRect(
      this.offsetLeft,
      this.height - this.style.scrollBarWidth,
      this.width - this.style.scrollBarWidth - this.offsetLeft,
      this.height
    );
    ctx.fillRect(
      this.width - this.style.scrollBarWidth,
      this.offsetTop,
      this.width,
      this.height - this.style.scrollBarWidth
    );
    ctx.fillStyle =
      this.state.isScrollXThumbHover || this.state.isSelectScrollXThumb
        ? this.style.scrollBarThumbActiveColor
        : this.style.scrollBarThumbColor;

    const scrollXThumbHeight = this.getScrollXThumbHeight();

    this.roundedRect(
      ctx,
      this.getScrollXThumbLeft(scrollXThumbHeight),
      this.height - this.style.scrollBarWidth + this.style.scrollBarThumbMargin,
      scrollXThumbHeight,
      this.style.scrollBarWidth - 2 * this.style.scrollBarThumbMargin,
      this.style.scrollBarThumbRadius
    );

    ctx.restore();
  }

  /** 绘制纵向滚动条 */
  drawScrollBarY(ctx: CanvasRenderingContext2D) {
    ctx.save();
    ctx.fillStyle = this.style.scrollBarBackgroundColor;
    ctx.strokeStyle = this.style.scrollBarBorderColor;
    ctx.lineWidth = this.style.scrollBarBorderWidth;
    ctx.fillRect(
      this.width - this.style.scrollBarWidth,
      this.height - this.style.scrollBarWidth,
      this.width,
      this.height
    );
    ctx.strokeRect(
      this.width - this.style.scrollBarWidth,
      this.offsetTop,
      this.width,
      this.height - this.style.scrollBarWidth - this.offsetTop
    );

    ctx.fillStyle =
      this.state.isScrollYThumbHover || this.state.isSelectScrollYThumb
        ? this.style.scrollBarThumbActiveColor
        : this.style.scrollBarThumbColor;

    const scrollYThumbHeight = this.getScrollYThumbHeight();
    this.roundedRect(
      ctx,
      this.width - this.style.scrollBarWidth + this.style.scrollBarThumbMargin,
      this.getScrollYThumbTop(scrollYThumbHeight),
      this.style.scrollBarWidth - 2 * this.style.scrollBarThumbMargin,
      scrollYThumbHeight,
      this.style.scrollBarThumbRadius
    );
    ctx.restore();
  }

  /** 绘制滚动条 */
  drawScrollBar(ctx: CanvasRenderingContext2D) {
    this.drawScrollBarX(ctx);
    this.drawScrollBarY(ctx);
  }

  /** 绘制带圆弧的长方形（绘制滚动条用） */
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

  /** 绘制标尺 */
  drawRuler(ctx: CanvasRenderingContext2D) {
    ctx.save();
    ctx.clearRect(0, 0, this.width, this.offsetTop);
    ctx.clearRect(0, 0, this.offsetLeft, this.height);
    const columns = this.viewCells[0].slice(1);
    const rows = this.viewCells.slice(1).map((cells) => cells[0]);
    ctx.textAlign = columns[0].style.textAlign as CanvasTextAlign;
    ctx.textBaseline = columns[0].style.textBaseline as CanvasTextBaseline;
    ctx.strokeStyle = this.style.rulerCellBorderColor;
    if (
      ctx.font !==
      `${columns[0].style.fontStyle} ${columns[0].style.fontWeight} ${columns[0].style.fontSize}pt ${columns[0].style.fontFamily}`
    ) {
      ctx.font = `${columns[0].style.fontStyle} ${columns[0].style.fontWeight} ${columns[0].style.fontSize}pt ${columns[0].style.fontFamily}`;
    }
    for (let len = columns.length, i = len - 1; i >= 0; i--) {
      if (columns[i].width === 0) {
        continue;
      }
      ctx.fillStyle = columns[i].style.background;
      const x =
        columns[i].x -
        this.scrollLeft +
        (columns[i - 1] && !columns[i - 1].width
          ? this.style.rulerResizeGapWidth / 2
          : 0);
      const y = columns[i].y;
      const width =
        columns[i].width -
        (columns[i + 1] && !columns[i + 1].width
          ? this.style.rulerResizeGapWidth / 2
          : 0) -
        (columns[i - 1] && !columns[i - 1].width
          ? this.style.rulerResizeGapWidth / 2
          : 0);
      const height = columns[i].height;
      let cellCtx: CanvasRenderingContext2D;
      const clip =
        columns[i].content.value &&
        ctx.measureText(columns[i].content.value).width > width;
      if (clip) {
        cellCtx = this.offscreenCtx;
        cellCtx.clearRect(0, 0, this.width, this.height);
        cellCtx.save();
        cellCtx.textAlign = columns[i].style.textAlign as CanvasTextAlign;
        cellCtx.textBaseline = columns[i].style
          .textBaseline as CanvasTextBaseline;
        cellCtx.strokeStyle = this.style.rulerCellBorderColor;
        cellCtx.font = ctx.font = `${columns[i].style.fontStyle} ${columns[i].style.fontWeight} ${columns[i].style.fontSize}pt ${columns[i].style.fontFamily}`;
        cellCtx.fillStyle = columns[i].style.background;
      } else {
        cellCtx = ctx;
        cellCtx.save();
      }
      cellCtx.fillRect(x, y, width, height);
      if (columns[i].content.value) {
        cellCtx.fillStyle = columns[0].style.color;
        const textMetrics = cellCtx.measureText(columns[i].content.value);
        cellCtx.fillText(
          columns[i].content.value,
          x + width / 2,
          columns[i].y +
            columns[i].height / 2 +
            ((textMetrics.actualBoundingBoxAscent || 4) -
              (textMetrics.actualBoundingBoxDescent || 0)) /
              2
        );
      }
      if (clip) {
        ctx.drawImage(
          cellCtx.canvas,
          x * this.multiple,
          y * this.multiple,
          width * this.multiple,
          height * this.multiple,
          x,
          y,
          width,
          height
        );
      }
      ctx.strokeRect(x + 0.5, y + 0.5, width, height);
      cellCtx.restore();
    }
    for (let len = rows.length, i = len - 1; i >= 0; i--) {
      if (rows[i].height === 0) {
        continue;
      }
      const x = rows[i].x;
      const y =
        rows[i].y -
        this.scrollTop +
        (rows[i - 1] && !rows[i - 1].height
          ? this.style.rulerResizeGapWidth / 2
          : 0);
      const width = rows[i].width;
      const height =
        rows[i].height -
        (rows[i + 1] && !rows[i + 1].height
          ? this.style.rulerResizeGapWidth / 2
          : 0) -
        (rows[i - 1] && !rows[i - 1].height
          ? this.style.rulerResizeGapWidth / 2
          : 0);
      let cellCtx: CanvasRenderingContext2D;
      const clip =
        rows[i].content.value &&
        rows[i].style.fontSize * 1.5 >
          rows[i].height - 2 * rows[i].style.borderWidth;
      if (clip) {
        cellCtx = this.offscreenCtx;
        cellCtx.clearRect(0, 0, this.width, this.height);
      } else {
        cellCtx = ctx;
      }
      cellCtx.save();
      cellCtx.textAlign = rows[i].style.textAlign as CanvasTextAlign;
      cellCtx.textBaseline = rows[i].style.textBaseline as CanvasTextBaseline;
      cellCtx.strokeStyle = this.style.rulerCellBorderColor;
      cellCtx.font = ctx.font = `${rows[i].style.fontStyle} ${rows[i].style.fontWeight} ${rows[i].style.fontSize}pt ${rows[i].style.fontFamily}`;
      // if (
      //   this.activeArr.length &&
      //   this.activeArr.find((range) =>
      //     inRange(rows[i].position.row, range.rowStart, range.rowEnd, true)
      //   )
      // ) {
      //   cellCtx.fillStyle = this.style.activeRulerCellBacgroundColor;
      // } else if (ctx.fillStyle !== this.style.rulerCellBackgroundColor) {
      //   cellCtx.fillStyle = this.style.rulerCellBackgroundColor;
      // }
      cellCtx.fillStyle = rows[i].style.background;
      cellCtx.fillRect(x, y, width, height);
      if (rows[i].content.value) {
        const textMetrics = cellCtx.measureText(rows[i].content.value);
        cellCtx.fillStyle = columns[0].style.color;
        cellCtx.fillText(
          rows[i].content.value,
          rows[i].x + rows[i].width / 2,
          y +
            height / 2 +
            ((textMetrics.actualBoundingBoxAscent || 4) -
              (textMetrics.actualBoundingBoxDescent || 0)) /
              2,
          rows[i].width - 2 * columns[0].style.borderWidth
        );
      }
      if (clip) {
        ctx.drawImage(
          cellCtx.canvas,
          rows[i].x * this.multiple,
          y * this.multiple,
          rows[i].width * this.multiple,
          height * this.multiple,
          rows[i].x,
          y,
          rows[i].width,
          height
        );
      }
      ctx.strokeRect(x + 0.5, y + 0.5, width, height);
      cellCtx.restore();
    }

    ctx.fillStyle = this.style.cellBackgroundColor;
    ctx.fillRect(0.5, 0.5, this.offsetLeft, this.offsetTop);
    ctx.strokeRect(0.5, 0.5, this.offsetLeft, this.offsetTop);
    ctx.fillStyle = this.style.activeRulerCellBacgroundColor;
    ctx.beginPath();
    ctx.moveTo(
      this.offsetLeft -
        this.style.rulerResizeGapWidth -
        this.style.allCellTriangleSideChief,
      this.offsetTop - this.style.rulerResizeGapWidth
    );
    ctx.lineTo(
      this.offsetLeft - this.style.rulerResizeGapWidth,
      this.offsetTop - this.style.rulerResizeGapWidth
    );
    ctx.lineTo(
      this.offsetLeft - this.style.rulerResizeGapWidth,
      this.offsetTop -
        this.style.rulerResizeGapWidth -
        this.style.allCellTriangleSideChief
    );
    ctx.fill();
    ctx.closePath();
    ctx.strokeRect(
      0.5,
      0.5,
      columns[columns.length - 1].x + columns[columns.length - 1].width,
      this.offsetTop
    );
    ctx.strokeRect(
      0.5,
      0.5,
      this.offsetLeft,
      rows[rows.length - 1].x + rows[rows.length - 1].height
    );
    ctx.restore();
  }

  /** 绘制视图中所有单元格 */
  drawPanel(ctx: CanvasRenderingContext2D) {
    for (let i = this.viewCells.length - 1; i > 0; i--) {
      for (let j = this.viewCells[i].length - 1; j > 0; j--) {
        this.drawCell(ctx, this.viewCells[i][j]);
      }
    }
  }

  drawFloat() {
    const ctx = this.offscreenCtx;
    ctx.clearRect(0, 0, this.width, this.height);
    this.floatCtx.clearRect(0, 0, this.width, this.height);
    const floatArr = this.filterOffscreenFloat(this.floatArr);
    if (!floatArr.length) {
      return;
    }
    ctx.save();
    for (const elem of floatArr) {
      ctx.drawImage(
        elem.content,
        elem.x - this.scrollLeft,
        elem.y - this.scrollTop,
        elem.width,
        elem.height
      );
    }
    this.floatCtx.drawImage(
      ctx.canvas,
      this.offsetLeft * this.multiple,
      this.offsetTop * this.multiple,
      this.clientWidth * this.multiple,
      this.clientHeight * this.multiple,
      this.offsetLeft,
      this.offsetTop,
      this.clientWidth,
      this.clientHeight
    );
    ctx.restore();
  }

  /** 绘制单个单元格 */
  drawCell(
    ctx: CanvasRenderingContext2D,
    drawCell: Cell,
    clip: boolean = false
  ) {
    const cell = drawCell.isCombined ? drawCell.combineCell : drawCell;
    const x =
      cell.x -
      (cell.type === 'rulerY' || cell.type === 'all  ' ? 0 : this.scrollLeft);
    const y =
      cell.y -
      (cell.type === 'rulerX' || cell.type === 'all' ? 0 : this.scrollTop);
    const width = cell.width;
    const height = cell.height;
    let cellCtx: CanvasRenderingContext2D;
    if (cell.content.html || cell.content.value) {
      const textArr = this.parseNode(
        this.htmlToElement(cell.content.html || cell.content.value)
      );
      const hasBold = !!textArr.find((text) => text.fontWeight === 'bold');
      const allHasItalic = textArr.every((text) => text.fontStyle === 'italic');
      const textMetricsArr = [];
      for (const textObj of textArr) {
        ctx.save();
        ctx.font = `${allHasItalic ? 'italic' : 'normal'} ${
          hasBold ? 'bold' : 'normal'
        } ${textObj.fontSize || cell.style.fontSize}pt ${
          textObj.fontFamily || cell.style.fontFamily
        }`;
        textMetricsArr.push(ctx.measureText(textObj.text));
        ctx.restore();
      }
      const maxTextHeight = Math.max(
        ...textMetricsArr.map(
          (metrics) =>
            metrics.actualBoundingBoxAscent +
              metrics.actualBoundingBoxDescent ||
            Math.max(
              ...textArr.map(
                (textObj) => textObj.fontSize || cell.style.fontSize
              )
            ) * 1.5
        )
      );
      const totalTextWidth = textMetricsArr
        .map((metrics) => metrics.width)
        .reduce((acc, cur) => acc + cur, 0);
      clip =
        clip ||
        totalTextWidth > cell.width - 2 * cell.style.borderWidth ||
        maxTextHeight > cell.height - 2 * cell.style.borderWidth;
    }
    if (clip) {
      cellCtx = this.offscreenCtx;
      cellCtx.clearRect(0, 0, this.width, this.height);
      cellCtx.save();
    } else {
      cellCtx = ctx;
    }
    if (cell.style.background) {
      if (cellCtx.fillStyle !== cell.style.background) {
        cellCtx.fillStyle = cell.style.background;
      }
      cellCtx.fillRect(x + 0.5, y + 0.5, width, height);
    }
    if (cellCtx.strokeStyle !== cell.style.borderColor) {
      cellCtx.strokeStyle = cell.style.borderColor;
    }
    cellCtx.strokeRect(x + 0.5, y + 0.5, width, height);
    if (cell.content.value || cell.content.html) {
      if (
        cellCtx.font !==
        `${cell.style.fontStyle} ${cell.style.fontWeight} ${cell.style.fontSize}pt ${cell.style.fontFamily}`
      ) {
        cellCtx.font = `${cell.style.fontStyle} ${cell.style.fontWeight} ${cell.style.fontSize}pt ${cell.style.fontFamily}`;
      }
      this.drawCellText(cell, x, y, width, height, cellCtx);
    }

    if (clip) {
      ctx.drawImage(
        cellCtx.canvas,
        x * this.multiple,
        y * this.multiple,
        width * this.multiple,
        height * this.multiple,
        x,
        y,
        width,
        height
      );
      cellCtx.restore();
    }
  }

  /** 绘制单元格显示文本 */
  drawCellText(
    cell: Cell,
    x: number,
    y: number,
    width: number,
    height: number,
    cellCtx: CanvasRenderingContext2D
  ) {
    if (cellCtx.fillStyle !== cell.style.color) {
      cellCtx.fillStyle = cell.style.color;
    }
    if (cell.style.textAlign && cellCtx.textAlign !== cell.style.textAlign) {
      cellCtx.textAlign = cell.style.textAlign as CanvasTextAlign;
    }
    if (
      cell.style.textBaseline &&
      cellCtx.textBaseline !== cell.style.textBaseline
    ) {
      cellCtx.textBaseline = cell.style.textBaseline as CanvasTextBaseline;
    }
    if (cell.content.html || cell.content.value) {
      const html = cell.content.html || cell.content.value;
      // console.log(this.htmlToElement(html).childNodes);
      // console.log(this.parseNode(this.htmlToElement(html)));
      const textArr = this.parseNode(this.htmlToElement(html));
      const hasBold = !!textArr.find((text) => text.fontWeight === 'bold');
      const allHasItalic = textArr.every((text) => text.fontStyle === 'italic');
      const textMetricsArr = [];
      for (const textObj of textArr) {
        cellCtx.save();
        cellCtx.font = `${textObj.fontStyle || cell.style.fontStyle} ${
          textObj.fontWeight || cell.style.fontWeight
        } ${textObj.fontSize || cell.style.fontSize}pt ${
          textObj.fontFamily || cell.style.fontFamily
        }`;
        textMetricsArr.push(cellCtx.measureText(textObj.text));
        cellCtx.restore();
      }
      const maxTextHeight = Math.max(
        ...textMetricsArr.map(
          (metrics) =>
            metrics.actualBoundingBoxAscent +
              metrics.actualBoundingBoxDescent || cell.style.fontSize * 1.5
        )
      );
      const totalTextWidth = textMetricsArr
        .map((metrics) => metrics.width)
        .reduce((acc, cur) => acc + cur, 0);
      textArr.forEach((textObj, index) => {
        cellCtx.save();
        if (cellCtx.fillStyle !== textObj.color) {
          cellCtx.fillStyle = textObj.color || cell.style.color;
        }
        if (
          (textObj.fontStyle && cell.style.fontStyle !== textObj.fontStyle) ||
          (textObj.fontFamily &&
            cell.style.fontFamily !== textObj.fontFamily) ||
          (textObj.fontSize && cell.style.fontSize !== textObj.fontSize) ||
          (textObj.fontWeight && cell.style.fontWeight !== textObj.fontWeight)
        ) {
          cellCtx.font = `${textObj.fontStyle || cell.style.fontStyle} ${
            textObj.fontWeight || cell.style.fontWeight
          } ${textObj.fontSize || cell.style.fontSize}pt ${
            textObj.fontFamily || cell.style.fontFamily
          }`;
        }

        let textX;
        let textY;
        if (cell.style.textAlign === this.style.cellTextAlignCenter) {
          textX =
            x +
            width / 2 -
            totalTextWidth / 2 +
            textMetricsArr
              .slice(0, index)
              .map((metrics) => metrics.width)
              .reduce((acc, cur) => acc + cur, 0) +
            textMetricsArr[index].width / 2 +
            0.5;
          const textHeight =
            textMetricsArr[index].actualBoundingBoxAscent +
              textMetricsArr[index].actualBoundingBoxDescent ||
            cell.style.fontSize * 1.5;
          textY =
            y +
            height / 2 +
            maxTextHeight / 2 -
            textHeight / 2 +
            ((textMetricsArr[index].actualBoundingBoxAscent || 4) -
              (textMetricsArr[index].actualBoundingBoxDescent || 0)) /
              2 +
            0.5;
        } else if (cell.style.textAlign === this.style.cellTextAlignLeft) {
          textX =
            x +
            cell.style.borderWidth +
            textMetricsArr
              .slice(0, index)
              .map((metrics) => metrics.width)
              .reduce((acc, cur) => acc + cur, 0) +
            0.5;
          const textHeight =
            textMetricsArr[index].actualBoundingBoxAscent +
              textMetricsArr[index].actualBoundingBoxDescent ||
            cell.style.fontSize;
          textY =
            y +
            height / 2 +
            maxTextHeight / 2 -
            textHeight / 2 +
            ((textMetricsArr[index].actualBoundingBoxAscent || 0) -
              (textMetricsArr[index].actualBoundingBoxDescent || 0)) /
              2 +
            0.5;
        } else if (cell.style.textAlign === this.style.cellTextAlignRight) {
          textX =
            x +
            width -
            cell.style.borderWidth -
            textMetricsArr
              .slice(index + 1)
              .map((metrics) => metrics.width)
              .reduce((acc, cur) => acc + cur, 0) +
            0.5;
          const textHeight =
            textMetricsArr[index].actualBoundingBoxAscent +
              textMetricsArr[index].actualBoundingBoxDescent ||
            cell.style.fontSize;
          textY =
            y +
            height / 2 +
            maxTextHeight / 2 -
            textHeight / 2 +
            ((textMetricsArr[index].actualBoundingBoxAscent || 0) -
              (textMetricsArr[index].actualBoundingBoxDescent || 0)) /
              2 +
            0.5;
        }
        cellCtx.fillText(textObj.text, textX, textY);

        cellCtx.restore();
      });
    }
    // const textMetrics = cellCtx.measureText(cell.content.value);
    // const textHeight =
    //   textMetrics.actualBoundingBoxAscent +
    //     textMetrics.actualBoundingBoxDescent || cell.style.fontSize * 1.5;
    // cellCtx.fillText(
    //   cell.content.value,
    //   x +
    //     (cell.style.textAlign === this.style.cellTextAlignCenter
    //       ? width / 2
    //       : cell.style.textAlign === this.style.cellTextAlignRight
    //       ? width - 2 * cell.style.borderWidth
    //       : cell.style.borderWidth) +
    //     0.5,
    //   y +
    //     height / 2 +
    //     ((textMetrics.actualBoundingBoxAscent || 4) -
    //       (textMetrics.actualBoundingBoxDescent || 0)) /
    //       2 +
    //     0.5
    // );
  }

  parseNode(parentNode, style: any = {}) {
    const res = [];
    if (parentNode.childNodes.length) {
      for (const node of parentNode.childNodes) {
        if (node instanceof Text) {
          res.push({
            text: node.textContent,
            textAlign: style.textAlign,
            color: style.color,
            fontSize:
              (style.fontSize && (parseInt(style.fontSize, 10) * 3) / 4) ||
              null,
            fontFamily: style.fontFamily,
            fontStyle: style.fontStyle,
            fontWeight: style.fontWeight,
            node,
            parentNode,
          });
        } else {
          res.push(
            ...this.parseNode(node, {
              textAlign: node.style.textAlign || style.textAlign,
              fontSize: node.style.fontSize || style.fontSize,
              color: node.style.color || style.color,
              fontFamily: node.style.fontFamily || style.fontFamily,
              fontStyle: node.style.fontStyle || style.fontStyle,
              fontWeight: node.style.fontWeight || style.fontWeight,
            })
          );
        }
      }
    }
    return res;
  }

  /** 绘制活动单元格or范围 */
  setActive() {
    this.actionCtx.clearRect(0, 0, this.width, this.height);
    this.floatActionCtx.clearRect(0, 0, this.width, this.height);
    const ctx = this.offscreenCtx;
    ctx.clearRect(0, 0, this.width, this.height);
    ctx.save();
    this.cells[0]
      .slice(1)
      .forEach(
        (cell) => (cell.style.background = this.style.rulerCellBackgroundColor)
      );
    this.cells
      .slice(1)
      .forEach(
        (row) => (row[0].style.background = this.style.rulerCellBackgroundColor)
      );
    const floatArr = this.filterOffscreenFloat(this.floatArr);
    if (floatArr.find((elem) => elem.isActive)) {
      ctx.lineWidth = this.style.activeFloatElementBorderWidth;
      ctx.strokeStyle = this.style.activeFloatElementBorderColor;
      ctx.fillStyle = '#fff';
      const radius = this.style.activeFloatElementResizeArcRadius;
      for (const elem of floatArr) {
        if (elem.isActive) {
          const x = elem.x - this.scrollLeft;
          const y = elem.y - this.scrollTop;
          const width = elem.width;
          const height = elem.height;
          ctx.strokeRect(x, y, width, height);
          ctx.beginPath();
          ctx.arc(x, y, radius, 0, 2 * Math.PI);
          ctx.fill();
          ctx.stroke();
          ctx.closePath();
          ctx.beginPath();
          ctx.arc(x + width / 2, y, radius, 0, 2 * Math.PI);
          ctx.fill();
          ctx.stroke();
          ctx.closePath();
          ctx.beginPath();
          ctx.arc(x + width, y, radius, 0, 2 * Math.PI);
          ctx.fill();
          ctx.stroke();
          ctx.closePath();
          ctx.beginPath();
          ctx.arc(x + width, y + height / 2, radius, 0, 2 * Math.PI);
          ctx.fill();
          ctx.stroke();
          ctx.closePath();
          ctx.beginPath();
          ctx.arc(x + width, y + height, radius, 0, 2 * Math.PI);
          ctx.fill();
          ctx.stroke();
          ctx.closePath();
          ctx.beginPath();
          ctx.arc(x + width / 2, y + height, radius, 0, 2 * Math.PI);
          ctx.fill();
          ctx.stroke();
          ctx.closePath();
          ctx.beginPath();
          ctx.arc(x, y + height, radius, 0, 2 * Math.PI);
          ctx.fill();
          ctx.stroke();
          ctx.closePath();
          ctx.beginPath();
          ctx.arc(x, y + height / 2, radius, 0, 2 * Math.PI);
          ctx.fill();
          ctx.stroke();
          ctx.closePath();
        }
      }
      this.floatActionCtx.drawImage(
        ctx.canvas,
        this.offsetLeft * this.multiple,
        this.offsetTop * this.multiple,
        this.clientWidth * this.multiple,
        this.clientHeight * this.multiple,
        this.offsetLeft,
        this.offsetTop,
        this.clientWidth,
        this.clientHeight
      );
    } else if (this.activeArr && this.activeArr.length) {
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
            Math.min(
              this.activeArr[i].columnStart,
              this.activeArr[i].columnEnd
            ),
            this.viewCells[0][1].position.column
          );
          const columnEnd = Math.min(
            Math.max(
              this.activeArr[i].columnStart,
              this.activeArr[i].columnEnd
            ),
            this.viewCells[0][this.viewCells[0].length - 1].position.column
          );

          this.cells[0].forEach((cell, index) => {
            if (index && inRange(index, columnStart, columnEnd, true)) {
              cell.style.background = this.style.activeRulerCellBacgroundColor;
            }
          });
          this.cells.forEach((row, index) => {
            if (index && inRange(index, rowStart, rowEnd, true)) {
              row[0].style.background = this.style.activeRulerCellBacgroundColor;
            }
          });

          ctx.fillStyle = this.style.selectedCellBackgroundColor;
          ctx.fillRect(
            this.cells[rowStart][columnStart].x -
              this.scrollLeft +
              2 * this.style.cellBorderWidth +
              0.5,
            this.cells[rowStart][columnStart].y -
              this.scrollTop +
              2 * this.style.cellBorderWidth +
              0.5,
            this.cells[rowEnd][columnEnd].x -
              this.cells[rowStart][columnStart].x +
              this.cells[rowEnd][columnEnd].width -
              4 * this.style.cellBorderWidth,
            this.cells[rowEnd][columnEnd].y -
              this.cells[rowStart][columnStart].y +
              this.cells[rowEnd][columnEnd].height -
              4 * this.style.cellBorderWidth
          );
          if (this.activeCellPos) {
            ctx.restore();
            ctx.strokeStyle = this.style.activeCellBorderColor;
            ctx.lineWidth = this.style.cellBorderWidth;
            const cell = this.cells[this.activeCellPos.row][
              this.activeCellPos.column
            ].isCombined
              ? this.cells[this.activeCellPos.row][this.activeCellPos.column]
                  .combineCell
              : this.cells[this.activeCellPos.row][this.activeCellPos.column];
            ctx.clearRect(
              cell.x - this.scrollLeft,
              cell.y - this.scrollTop,
              cell.width,
              cell.height
            );
            ctx.strokeRect(
              cell.x - this.scrollLeft + 2 * this.style.cellBorderWidth + 0.5,
              cell.y - this.scrollTop + 2 * this.style.cellBorderWidth + 0.5,
              cell.width - 4 * this.style.cellBorderWidth,
              cell.height - 4 * this.style.cellBorderWidth
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

        this.cells[0].forEach((cell, index) => {
          if (index && inRange(index, columnStart, columnEnd, true)) {
            cell.style.background = this.style.activeRulerCellBacgroundColor;
          }
        });
        this.cells.forEach((row, index) => {
          if (index && inRange(index, rowStart, rowEnd, true)) {
            row[0].style.background = this.style.activeRulerCellBacgroundColor;
          }
        });

        ctx.fillStyle = this.style.selectedCellBackgroundColor;
        ctx.fillRect(
          this.cells[rowStart][columnStart].x - this.scrollLeft + 0.5,
          this.cells[rowStart][columnStart].y - this.scrollTop + 0.5,
          this.cells[rowEnd][columnEnd].x -
            this.cells[rowStart][columnStart].x +
            this.cells[rowEnd][columnEnd].width,
          this.cells[rowEnd][columnEnd].y -
            this.cells[rowStart][columnStart].y +
            this.cells[rowEnd][columnEnd].height
        );
        if (this.activeCellPos) {
          const cell = this.cells[this.activeCellPos.row][
            this.activeCellPos.column
          ].isCombined
            ? this.cells[this.activeCellPos.row][this.activeCellPos.column]
                .combineCell
            : this.cells[this.activeCellPos.row][this.activeCellPos.column];
          ctx.clearRect(
            cell.x - this.scrollLeft + 0.5,
            cell.y - this.scrollTop + 0.5,
            cell.width,
            cell.height
          );
        }
        ctx.strokeStyle = this.style.activeCellBorderColor;
        ctx.lineWidth = this.style.cellBorderWidth;
        // ctx.shadowColor = this.style.activeCellShadowColor;
        // ctx.shadowBlur = this.style.activeCellShadowBlur;
        ctx.strokeRect(
          this.cells[rowStart][columnStart].x - this.scrollLeft + 0.5,
          this.cells[rowStart][columnStart].y - this.scrollTop + 0.5,
          this.cells[rowEnd][columnEnd].x -
            this.cells[rowStart][columnStart].x +
            this.cells[rowEnd][columnEnd].width,
          this.cells[rowEnd][columnEnd].y -
            this.cells[rowStart][columnStart].y +
            this.cells[rowEnd][columnEnd].height
        );
      }
      ctx.restore();
      if (this.unActiveRange) {
        ctx.save();
        const rowStart = Math.max(
          Math.min(this.unActiveRange.rowEnd, this.unActiveRange.rowStart),
          this.viewCells[1][0].position.row
        );
        const rowEnd = Math.min(
          Math.max(this.unActiveRange.rowEnd, this.unActiveRange.rowStart),
          this.viewCells[this.viewCells.length - 1][0].position.row
        );
        const columnStart = Math.max(
          Math.min(
            this.unActiveRange.columnStart,
            this.unActiveRange.columnEnd
          ),
          this.viewCells[0][1].position.column
        );
        const columnEnd = Math.min(
          Math.max(
            this.unActiveRange.columnStart,
            this.unActiveRange.columnEnd
          ),
          this.viewCells[0][this.viewCells[0].length - 1].position.column
        );
        ctx.strokeStyle = this.style.cellBorderColor;
        ctx.lineWidth = 2 * this.style.cellBorderWidth;
        ctx.fillStyle = this.style.unSelectedCellBackgroundColor;
        ctx.beginPath();
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
        ctx.closePath();
      }
      this.actionCtx.drawImage(
        ctx.canvas,
        this.offsetLeft * this.multiple,
        this.offsetTop * this.multiple,
        this.clientWidth * this.multiple,
        this.clientHeight * this.multiple,
        this.offsetLeft,
        this.offsetTop,
        this.clientWidth,
        this.clientHeight
      );
    }
    ctx.restore();
  }

  /** 绘制剪切状态边框 */
  setClipStatusAnimation(offset = 0) {
    if (this.clipAnimationTimeoutID) {
      clearTimeout(this.clipAnimationTimeoutID);
    }
    this.animationCtx.clearRect(0, 0, this.width, this.height);
    if (this.clipBoard) {
      const ctx = this.offscreenCtx;
      ctx.clearRect(0, 0, this.width, this.height);
      ctx.save();
      const rowStart = Math.max(
        Math.min(this.clipBoard.range.rowEnd, this.clipBoard.range.rowStart),
        this.viewCells[1][0].position.row
      );
      const rowEnd = Math.min(
        Math.max(this.clipBoard.range.rowEnd, this.clipBoard.range.rowStart),
        this.viewCells[this.viewCells.length - 1][0].position.row
      );
      const columnStart = Math.max(
        Math.min(
          this.clipBoard.range.columnStart,
          this.clipBoard.range.columnEnd
        ),
        this.viewCells[0][1].position.column
      );
      const columnEnd = Math.min(
        Math.max(
          this.clipBoard.range.columnStart,
          this.clipBoard.range.columnEnd
        ),
        this.viewCells[0][this.viewCells[0].length - 1].position.column
      );
      ctx.strokeStyle = this.style.activeCellBorderColor;
      ctx.lineWidth = this.style.cellBorderWidth;
      ctx.setLineDash([8, 2]);
      ctx.lineDashOffset = -1 * offset;
      ctx.strokeRect(
        this.cells[rowStart][columnStart].x -
          this.scrollLeft +
          this.style.cellBorderWidth,
        this.cells[rowStart][columnStart].y -
          this.scrollTop +
          this.style.cellBorderWidth,
        this.cells[rowEnd][columnEnd].x -
          this.cells[rowStart][columnStart].x +
          this.cells[rowEnd][columnEnd].width -
          2 * this.style.cellBorderWidth,
        this.cells[rowEnd][columnEnd].y -
          this.cells[rowStart][columnStart].y +
          this.cells[rowEnd][columnEnd].height -
          2 * this.style.cellBorderWidth
      );

      this.animationCtx.drawImage(
        ctx.canvas,
        this.offsetLeft * this.multiple,
        this.offsetTop * this.multiple,
        this.clientWidth * this.multiple,
        this.clientHeight * this.multiple,
        this.offsetLeft,
        this.offsetTop,
        this.clientWidth,
        this.clientHeight
      );
      this.clipAnimationTimeoutID = setTimeout(() => {
        this.setClipStatusAnimation((offset + 2 < 16 && offset + 2) || 0);
      }, 100);
      ctx.restore();
    } else {
      this.clipAnimationTimeoutID = null;
    }
  }

  /** 生成列标尺内容 */
  generateColumnNum(columnIndex: number): string {
    const columnNumArr: string[] = [];
    while (columnIndex > 26) {
      columnNumArr.unshift(String.fromCharCode((columnIndex % 26 || 26) + 64));
      columnIndex = Math.floor(columnIndex / 26);
    }
    columnNumArr.unshift(String.fromCharCode(columnIndex + 64));

    return columnNumArr.join('');
  }

  /** 生成行标尺内容 */
  generateRowNum(rowIndex: number): string {
    return rowIndex > 0 ? rowIndex.toString() : '';
  }

  onMouseOver(event: MouseEvent) {
    // console.log('mouseover', event);
  }

  /** 鼠标移出绘制区域事件 */
  onMouseOut(event: MouseEvent) {
    // console.log('mouseout', event);
    // this.state.isSelectCell = false;
    // this.state.isScrollXThumbHover = false;
    // this.state.isScrollYThumbHover = false;
    // this.state.isSelectScrollXThumb = false;
    // this.state.isSelectScrollYThumb = false;
    // this.state.isSelectRulerX = false;
    // this.state.isSelectRulerY = false;
    // this.state.isResizeColumn = false;
    // this.state.isResizeRow = false;
    // this.resizeColumnCell = null;
    // this.resizeRowCell = null;
    // this.drawScrollBar(this.ctx);
    // // this.mousePoint = null;
    // if (
    //   this.state.unSelectCell ||
    //   this.state.unSelectRulerX ||
    //   this.state.unSelectRulerY
    // ) {
    //   this.reCalcActiveArr();
    //   this.unActiveRange = null;
    //   this.state.unSelectCell = false;
    //   this.state.unSelectRulerX = false;
    //   this.state.unSelectRulerY = false;
    //   this.setActive();
    //   this.drawScrollBar(this.ctx);
    //   this.drawRuler(this.ctx);
    // }
  }

  onMouseEnter(event) {
    // console.log('mouseenter', event);
  }

  onMouseLeave(event) {
    // console.log('mouseleave', event);
  }

  /** 鼠标右键事件 */
  onContextMenu(event: MouseEvent) {
    event.returnValue = false;
  }

  /** 判断坐标点是否在单元格中 */
  inCellArea(x: number, y: number, cell: Cell) {
    return (
      !cell.isCombined &&
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

  inRect(x: number, y: number, rx: number, ry: number, rw: number, rh: number) {
    return inRange(x, rx, rx + rw) && inRange(y, ry, ry + rh);
  }

  /** 判断坐标点是否在非标尺单元格区域 */
  inCellsArea(x: number, y: number) {
    return (
      inRange(x, this.offsetLeft, this.offsetLeft + this.clientWidth, true) &&
      inRange(y, this.offsetTop, this.offsetTop + this.clientHeight, true)
    );
  }

  /** 判断坐标点是否在列标尺区域 */
  inRulerXArea(x: number, y: number) {
    return (
      inRange(x, this.offsetLeft, this.width, true) &&
      inRange(y, 0, this.offsetTop, true)
    );
  }

  /** 判断坐标点是否在列可调整区域 */
  inRulerXResizeGap(x: number, y: number) {
    const columns = this.viewCells[0];
    if (this.inRulerXArea(x, y)) {
      for (let i = 1, len = columns.length; i < len; i++) {
        if (
          inRange(
            x,
            columns[i].x -
              this.scrollLeft +
              columns[i].width -
              this.style.rulerResizeGapWidth,
            columns[i].x -
              this.scrollLeft +
              columns[i].width +
              this.style.rulerResizeGapWidth
          )
        ) {
          return true;
        }
      }
    }
    return false;
  }

  /** 判断坐标是否在行调整区域 */
  inRulerYResizeGap(x: number, y: number) {
    const rows = this.viewCells.map((cells) => cells[0]);
    if (this.inRulerYArea(x, y)) {
      for (let i = 1, len = rows.length; i < len; i++) {
        if (
          inRange(
            y,
            rows[i].y -
              this.scrollTop +
              rows[i].height -
              this.style.rulerResizeGapWidth,
            rows[i].y -
              this.scrollTop +
              rows[i].height +
              this.style.rulerResizeGapWidth
          )
        ) {
          return true;
        }
      }
    }
    return false;
  }

  /** 判断是否在行标尺区域 */
  inRulerYArea(x: number, y: number) {
    return (
      inRange(x, 0, this.offsetLeft, true) &&
      inRange(y, this.offsetTop, this.height, true)
    );
  }

  /** 判断是否在纵向滚动绘制区域 */
  inScrollYBarArea(x: number, y: number) {
    return (
      inRange(x, this.offsetLeft + this.clientWidth, this.width, true) &&
      inRange(y, this.offsetTop, this.offsetTop + this.clientHeight, true)
    );
  }

  /** 判断是否在横向滚动绘制区域 */
  inScrollXBarArea(x: number, y: number) {
    return (
      inRange(x, this.offsetLeft, this.width, true) &&
      inRange(y, this.offsetTop + this.clientHeight, this.height, true)
    );
  }

  /** 判断是否在横向滚动条上 */
  inThumbAreaOfScrollBarX(x: number, y: number, judegedInScrollBarX = false) {
    if (!judegedInScrollBarX && !this.inScrollXBarArea(x, y)) {
      return;
    }
    const scrollXThumbHeight = this.getScrollXThumbHeight();
    const scrollXThumbLeft = this.getScrollXThumbLeft(scrollXThumbHeight);

    return inRange(x, scrollXThumbLeft, scrollXThumbLeft + scrollXThumbHeight);
  }

  /** 获取横向滚动条宽度 */
  getScrollXThumbHeight() {
    let scrollXThumbHeight =
      (this.clientWidth / this.scrollWidth) * this.clientWidth;
    if (scrollXThumbHeight < this.style.scrollBarThumbMinSize) {
      scrollXThumbHeight = this.style.scrollBarThumbMinSize;
    }
    return scrollXThumbHeight;
  }

  /** 获取横向滚动条距左侧距离 */
  getScrollXThumbLeft(scrollXThumbHeight: number) {
    return (
      this.offsetLeft +
      (this.scrollLeft * (this.clientWidth - scrollXThumbHeight)) /
        (this.scrollWidth - this.clientWidth)
    );
  }

  /** 判断是否在纵向滚动条上 */
  inThumbAreaOfScrollBarY(x: number, y: number, judegedInScrollBarY = false) {
    if (!judegedInScrollBarY && !this.inScrollYBarArea(x, y)) {
      return false;
    }
    const scrollYThumbHeight = this.getScrollYThumbHeight();
    const scrollYThumbTop = this.getScrollYThumbTop(scrollYThumbHeight);

    return inRange(y, scrollYThumbTop, scrollYThumbTop + scrollYThumbHeight);
  }

  /** 获取纵向滚动条高度 */
  getScrollYThumbHeight() {
    let scrollYThumbHeight =
      (this.clientHeight / this.scrollHeight) * this.clientHeight;
    if (scrollYThumbHeight < this.style.scrollBarThumbMinSize) {
      scrollYThumbHeight = this.style.scrollBarThumbMinSize;
    }
    return scrollYThumbHeight;
  }

  /** 获取纵向滚动条距上侧距离 */
  getScrollYThumbTop(scrollYThumbHeight: number) {
    return (
      this.offsetTop +
      (this.scrollTop * (this.clientHeight - scrollYThumbHeight)) /
        (this.scrollHeight - this.clientHeight)
    );
  }

  /** 判断是否在全选区域 */
  inSelectAllArea(x: number, y: number) {
    return inRange(x, 0, this.offsetLeft) && inRange(y, 0, this.offsetTop);
  }

  /** 鼠标点下时事件 */
  onMouseDown(event: MouseEvent) {
    // this.panel.nativeElement.focus();
    // const eventX = event.offsetX * this.multiple;
    // const eventY = event.offsetY * this.multiple;
    const eventX = event.offsetX / this.multiple;
    const eventY = event.offsetY / this.multiple;
    event.preventDefault();
    event.stopPropagation();
    this.canvas.focus();
    // console.log('mousedown', event);

    if (this.state.isCellEdit || this.editingCell) {
      this.editCellCompelte();
    }

    if (event.button === 2) {
      event.returnValue = false;
      return;
    }
    const currentTime = new Date().getTime();
    let isDblClick = false;
    if (
      this.mousePoint &&
      this.mousePoint.lastModifyTime &&
      currentTime - this.mousePoint.lastModifyTime < 300
    ) {
      isDblClick = true;
    }
    this.mousePoint = {
      x: eventX,
      y: eventY,
      lastModifyTime: new Date().getTime(),
    };
    if (this.inSelectAllArea(eventX, eventY)) {
      // console.log('all');
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
    } else if (this.inRulerXArea(eventX, eventY)) {
      // console.log('rulerx');
      for (let i = 1, len = this.viewCells[0].length - 1; i < len; i++) {
        if (
          this.viewCells[0][i].width >
            (i === 1 ? 1 : 2) * this.style.rulerResizeGapWidth &&
          inRange(
            eventX,
            this.viewCells[0][i].x -
              this.scrollLeft +
              (i === 1 ? 0 : this.style.rulerResizeGapWidth),
            this.viewCells[0][i].x -
              this.scrollLeft +
              this.viewCells[0][i].width -
              this.style.rulerResizeGapWidth,
            true
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
            eventX,
            this.viewCells[0][i].x +
              this.viewCells[0][i].width -
              this.style.rulerResizeGapWidth <
              this.viewCells[0][i].x
              ? this.viewCells[0][i].x - this.scrollLeft
              : this.viewCells[0][i].x -
                  this.scrollLeft +
                  this.viewCells[0][i].width -
                  this.style.rulerResizeGapWidth,
            this.viewCells[0][i].x +
              this.viewCells[0][i].width +
              this.style.rulerResizeGapWidth >
              this.viewCells[0][i + 1].x + this.viewCells[0][i + 1].width
              ? this.viewCells[0][i + 1].width
                ? this.viewCells[0][i + 1].x + this.viewCells[0][i + 1].width
                : this.viewCells[0][i].x +
                  this.viewCells[0][i].width -
                  this.scrollLeft
              : this.viewCells[0][i].x -
                  this.scrollLeft +
                  this.viewCells[0][i].width +
                  this.style.rulerResizeGapWidth,
            true
          )
        ) {
          if (isDblClick) {
            const width = this.cells
              .map((row) => row[i])
              .reduce((acc, cur) => {
                if (cur.content.value) {
                  this.ctx.save();
                  this.ctx.font = `${cur.style.fontStyle} ${cur.style.fontWeight} ${cur.style.fontSize}pt ${cur.style.fontFamily}`;
                  if (
                    this.ctx.measureText(cur.content.value).width +
                      2 * this.style.cellBorderWidth >
                    acc
                  ) {
                    return (
                      this.ctx.measureText(cur.content.value).width +
                      2 * this.style.cellBorderWidth
                    );
                  }
                }
                return acc;
              }, 0);
            if (width) {
              this.resizeColumn(
                this.viewCells[0][i].position.column,
                width - this.viewCells[0][i].width
              );
            }
          } else {
            this.resizeColumnCell = this.viewCells[0][i];
            this.state.isResizeColumn = true;
          }
          break;
        }
      }
    } else if (this.inRulerYArea(eventX, eventY)) {
      // console.log('rulery', event);
      const rowCells = this.viewCells.map((row) => row[0]);
      for (let i = 1, len = rowCells.length - 1; i < len; i++) {
        if (
          rowCells[i].height >
            (i === 1 ? 1 : 2) * this.style.rulerResizeGapWidth &&
          inRange(
            eventY,
            rowCells[i].y -
              this.scrollTop +
              (i === 1 ? 0 : this.style.rulerResizeGapWidth),
            rowCells[i].y -
              this.scrollTop +
              rowCells[i].height -
              this.style.rulerResizeGapWidth,
            true
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
            eventY,
            rowCells[i].y +
              rowCells[i].height -
              this.style.rulerResizeGapWidth <
              rowCells[i].y
              ? rowCells[i].y - this.scrollTop
              : rowCells[i].y -
                  this.scrollTop +
                  rowCells[i].height -
                  this.style.rulerResizeGapWidth,
            rowCells[i].y +
              rowCells[i].height +
              this.style.rulerResizeGapWidth >
              rowCells[i + 1].y + rowCells[i + 1].height
              ? rowCells[i + 1].height
                ? rowCells[i + 1].y + rowCells[i + 1].height - this.scrollTop
                : rowCells[i].y + rowCells[i].height - this.scrollTop
              : rowCells[i].y -
                  this.scrollTop +
                  rowCells[i].height +
                  this.style.rulerResizeGapWidth
          )
        ) {
          if (isDblClick) {
            this.resizeRow(
              rowCells[i].position.row,
              this.style.cellHeight - rowCells[i].height
            );
            break;
          } else {
            this.resizeRowCell = rowCells[i];
            this.state.isResizeRow = true;
            break;
          }
        }
      }
    } else if (this.inScrollXBarArea(eventX, eventY)) {
      // console.log('scrollx');
      if (this.inThumbAreaOfScrollBarX(eventX, eventY, true)) {
        this.state.isSelectScrollXThumb = true;
      } else if (
        inRange(
          eventX,
          this.offsetLeft,
          this.getScrollXThumbLeft(this.getScrollXThumbHeight())
        )
      ) {
        this.scrollX((-1 * this.viewColumnCount * this.style.cellWidth) / 2);
        this.setActive();
        this.drawRuler(this.ctx);
      } else {
        this.scrollX((this.viewColumnCount * this.style.cellWidth) / 2);
        this.setActive();
        this.drawRuler(this.ctx);
      }
    } else if (this.inScrollYBarArea(eventX, eventY)) {
      // console.log('scrolly');
      if (this.inThumbAreaOfScrollBarY(eventX, eventY, true)) {
        this.state.isSelectScrollYThumb = true;
      } else if (
        inRange(
          eventY,
          this.offsetTop,
          this.getScrollYThumbTop(this.getScrollYThumbHeight())
        )
      ) {
        this.scrollY((-1 * this.viewRowCount * this.style.cellHeight) / 2);
        this.setActive();
        this.drawRuler(this.ctx);
      } else {
        this.scrollY((this.viewRowCount * this.style.cellHeight) / 2);
        this.setActive();
        this.drawRuler(this.ctx);
      }
    } else if (this.inCellsArea(eventX, eventY)) {
      // for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
      //   for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
      //     const cell = this.viewCells[i][j].isCombined
      //       ? this.viewCells[i][j].combineCell
      //       : this.viewCells[i][j];
      //     if (this.inCellArea(eventX, eventY, cell)) {
      const floatElement = this.getViewFloatElementByPoint(eventX, eventY);

      if (floatElement) {
        const floatElemntResizePos = this.getFloatElementResizePos(
          floatElement,
          eventX,
          eventY
        );
        if (
          floatElemntResizePos !== LogicPosition.Other &&
          floatElement.isActive
        ) {
          this.state.isResizeFloat = true;
          this.state.resizeFloatPos = floatElemntResizePos;
        } else {
          this.state.isMoveFloat = true;
          floatElement.moveOrigin = {
            x: eventX + this.scrollLeft - floatElement.x,
            y: eventY + this.scrollTop - floatElement.y,
          };
        }
        if (event.ctrlKey || event.shiftKey) {
          floatElement.isActive = true;
        } else {
          this.floatArr.forEach(
            (elem) => (elem.isActive = floatElement === elem)
          );
        }

        this.activeCellPos = null;
        this.activeArr = [];
      } else {
        this.floatArr.forEach(
          (elem) => (elem.isActive = floatElement === elem)
        );
        const cell = this.getViewCellByPoint(eventX, eventY);
        if (!cell) {
          return;
        }
        const isUnActive =
          this.activeArr.some(
            (range) =>
              inRange(cell.position.row, range.rowStart, range.rowEnd, true) &&
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
            rowEnd: cell.position.row + cell.rowSpan - 1,
            columnStart: cell.position.column,
            columnEnd: cell.position.column + cell.colSpan - 1,
          };
          this.unActiveCellPos = {
            row: cell.position.row,
            column: cell.position.column,
          };
        } else {
          if (
            isDblClick &&
            cell.position.row === this.activeCellPos.row &&
            cell.position.column === this.activeCellPos.column
          ) {
            this.resetCellPerspective(cell);
            this.clipBoard = null;
            this.editingCell = cell;
            this.editingCell.content.previousValue = this.editingCell.content.value;
            this.editingCell.content.previousHtml = this.editingCell.content.html;
            this.state.isCellEdit = true;
            this.activeArr = [
              {
                rowStart: cell.position.row,
                rowEnd: cell.position.row + cell.rowSpan - 1,
                columnStart: cell.position.column,
                columnEnd: cell.position.column + cell.colSpan - 1,
              },
            ];
            this.activeCellPos.rangeIndex = 0;
          } else {
            this.state.isSelectCell = true;
            if (event.shiftKey) {
              const range = {
                rowStart: this.activeCellPos.row,
                rowEnd: cell.position.row,
                columnStart: this.activeCellPos.column,
                columnEnd: cell.position.column,
              };
              this.recalcRange(range);
              this.activeArr.push(range);
            } else {
              this.activeArr.push({
                rowStart: cell.position.row,
                columnStart: cell.position.column,
                rowEnd: cell.position.row + cell.rowSpan - 1,
                columnEnd: cell.position.column + cell.colSpan - 1,
              });
              this.activeCellPos = {
                row: cell.position.row,
                column: cell.position.column,
                rangeIndex: this.activeArr.length - 1,
              };
            }
            // this.activeArr.push({
            //   rowStart:
            //     (event.shiftKey &&
            //       this.activeArr.length &&
            //       this.activeArr[this.activeArr.length - 1].rowStart) ||
            //     cell.position.row,
            //   columnStart:
            //     (event.shiftKey &&
            //       this.activeArr.length &&
            //       this.activeArr[this.activeArr.length - 1].columnStart) ||
            //     cell.position.column,
            //   rowEnd: cell.position.row + cell.rowSpan - 1,
            //   columnEnd: cell.position.column + cell.colSpan - 1,
            // });
            // this.activeCellPos = {
            //   row: this.activeArr[this.activeArr.length - 1].rowStart,
            //   column: this.activeArr[this.activeArr.length - 1].columnStart,
            //   rangeIndex: this.activeArr.length - 1,
            // };
            if (!event.ctrlKey || event.shiftKey) {
              this.activeArr = [this.activeArr[this.activeArr.length - 1]];
              this.activeCellPos.rangeIndex = 0;
            }
          }
        }
        this.autoScroll(this.mousePoint.x, this.mousePoint.y);
      }

      this.setActive();
      this.drawRuler(this.ctx);
      return;
      //       }
      //     }
      //   }
    }
  }

  // 设置鼠标cursor样式
  setMouseCursor(eventX, eventY) {
    if (
      (!this.inCellsArea(eventX, eventY) && !this.state.isSelectCell) ||
      this.state.isSelectScrollYThumb ||
      this.state.isSelectScrollXThumb ||
      this.state.isResizeColumn ||
      this.state.isResizeRow
    ) {
      if (
        this.state.isResizeColumn ||
        (!this.state.isSelectRulerX &&
          !this.state.isSelectRulerY &&
          this.inRulerXResizeGap(eventX, eventY))
      ) {
        this.canvas.style.cursor = 'col-resize';
      } else if (
        this.state.isResizeRow ||
        (!this.state.isSelectRulerX &&
          !this.state.isSelectRulerY &&
          this.inRulerYResizeGap(eventX, eventY))
      ) {
        this.canvas.style.cursor = 'row-resize';
      } else {
        this.canvas.style.cursor = 'default';
      }
    } else if (this.state.isMoveFloat) {
      this.canvas.style.cursor = 'move';
    } else {
      const floatElement = this.getViewFloatElementByPoint(eventX, eventY);
      if (floatElement || this.state.isResizeFloat) {
        const pos =
          this.state.resizeFloatPos ||
          this.getFloatElementResizePos(floatElement, eventX, eventY);
        switch (pos) {
          case LogicPosition.LeftTop:
          case LogicPosition.RightBottom:
            this.canvas.style.cursor = 'nwse-resize';
            break;
          case LogicPosition.rightTop:
          case LogicPosition.LeftBottom:
            this.canvas.style.cursor = 'nesw-resize';
            break;
          case LogicPosition.Top:
          case LogicPosition.Bottom:
            this.canvas.style.cursor = 'ns-resize';
            break;
          case LogicPosition.Left:
          case LogicPosition.Right:
            this.canvas.style.cursor = 'ew-resize';
            break;
          default:
            this.canvas.style.cursor = 'move';
            break;
        }
      } else {
        this.canvas.style.cursor = 'cell';
      }
    }
  }

  // @throttle(20)
  /** 鼠标移动事件 */
  onMouseMove(event: any) {
    // console.log('move');
    const eventX = (event.offsetX || event.layerX || 0) / this.multiple; // 兼容火狐读offetX，offsetY一直为0问题
    const eventY = (event.offsetY || event.layerY || 0) / this.multiple;
    this.setMouseCursor(eventX, eventY);
    const preIsScrollXThumbHover = this.state.isScrollXThumbHover;
    const preIsScrollYThumbHover = this.state.isScrollYThumbHover;
    this.state.isScrollXThumbHover = this.inThumbAreaOfScrollBarX(
      eventX,
      eventY
    );
    this.state.isScrollYThumbHover = this.inThumbAreaOfScrollBarY(
      eventX,
      eventY
    );
    if (
      (preIsScrollXThumbHover !== this.state.isScrollXThumbHover &&
        !this.state.isSelectScrollXThumb) ||
      (preIsScrollYThumbHover !== this.state.isScrollYThumbHover &&
        !this.state.isSelectScrollYThumb)
    ) {
      this.drawScrollBar(this.ctx);
    }

    // if (!this.isTicking) {
    //   requestAnimationFrame(() => {
    if (this.state.isMoveFloat) {
      this.moveFloat(eventX, eventY);
    } else if (this.state.isResizeFloat) {
      this.resizeFloat(eventX, eventY);
    } else {
      let range;
      if (this.state.isSelectCell) {
        range = this.calcActive(eventX, eventY);
      } else if (this.state.unSelectCell) {
        range = this.calcUnActive(eventX, eventY);
      } else if (this.state.isSelectScrollYThumb) {
        this.calcScrollY(eventX, eventY);
      } else if (this.state.isSelectScrollXThumb) {
        this.calcScrollX(eventX, eventY);
      } else if (this.state.isSelectRulerX) {
        range = this.calcAcitiveRulerX(eventX, eventY);
      } else if (this.state.isSelectRulerY) {
        range = this.calcActiveRulerY(eventX, eventY);
      } else if (this.state.unSelectRulerX) {
        range = this.calcUnActiveRulerX(eventX, eventY);
      } else if (this.state.unSelectRulerY) {
        range = this.calcUnActiveRulerY(eventX, eventY);
      } else if (this.state.isResizeColumn) {
        this.resizeColumn(
          this.resizeColumnCell.position.column,
          eventX - this.mousePoint.x
        );
      } else if (this.state.isResizeRow) {
        this.resizeRow(
          this.resizeRowCell.position.row,
          eventY - this.mousePoint.y
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
    }
    this.mousePoint = {
      x: eventX,
      y: eventY,
      lastModifyTime: this.mousePoint && this.mousePoint.lastModifyTime,
    };
    //   this.isTicking = false;
    // });
    // }

    // this.isTicking = true;
  }

  resizeFloat(eventX, eventY) {
    const activeArr = this.floatArr.filter((elem) => elem.isActive);
    for (const floatElem of activeArr) {
      switch (this.state.resizeFloatPos) {
        case LogicPosition.LeftTop:
          floatElem.x = eventX + this.scrollLeft;
          floatElem.y = eventY + this.scrollTop;
          floatElem.width -= eventX - this.mousePoint.x;
          floatElem.height -= eventY - this.mousePoint.y;
          break;
        case LogicPosition.LeftBottom:
          floatElem.x = eventX + this.scrollLeft;
          floatElem.width -= eventX - this.mousePoint.x;
          floatElem.height += eventY - this.mousePoint.y;
          break;
        case LogicPosition.Left:
          floatElem.x = eventX + this.scrollLeft;
          floatElem.width -= eventX - this.mousePoint.x;
          break;
        case LogicPosition.Right:
          floatElem.width += eventX - this.mousePoint.x;
          break;
        case LogicPosition.Top:
          floatElem.y = eventY + this.scrollTop;
          floatElem.height -= eventY - this.mousePoint.y;
          break;
        case LogicPosition.Bottom:
          floatElem.height += eventY - this.mousePoint.y;
          break;
        case LogicPosition.rightTop:
          floatElem.y = eventY + this.scrollTop;
          floatElem.width += eventX - this.mousePoint.x;
          floatElem.height -= eventY - this.mousePoint.y;
          break;
        case LogicPosition.RightBottom:
          floatElem.width += eventX - this.mousePoint.x;
          floatElem.height += eventY - this.mousePoint.y;
          break;
      }
      floatElem.x =
        floatElem.x < this.offsetLeft ? this.offsetLeft : floatElem.x;
      floatElem.y = floatElem.y < this.offsetTop ? this.offsetTop : floatElem.y;
      floatElem.width = floatElem.width < 0 ? 0 : floatElem.width;
      floatElem.height = floatElem.height < 0 ? 0 : floatElem.height;
    }
    this.refreshView();
  }

  /** 移动悬浮元素（图片） */
  moveFloat(eventX, eventY) {
    const activeArr = this.floatArr.filter((elem) => elem.isActive);
    for (const floatElem of activeArr) {
      if (
        !(
          eventX + this.scrollLeft - this.offsetLeft < floatElem.moveOrigin.x &&
          eventX > this.mousePoint.x
        )
      ) {
        floatElem.x += eventX - this.mousePoint.x;
      }
      if (
        !(
          eventY + this.scrollTop - this.offsetTop < floatElem.moveOrigin.y &&
          eventY > this.mousePoint.y
        )
      ) {
        floatElem.y += eventY - this.mousePoint.y;
      }

      floatElem.x =
        floatElem.x < this.offsetLeft ? this.offsetLeft : floatElem.x;
      floatElem.y = floatElem.y < this.offsetTop ? this.offsetTop : floatElem.y;
    }

    this.refreshView();
  }

  /** 重新设置列宽度 */
  resizeColumn(column: number, deltaWidth: number) {
    if (deltaWidth < -1 * this.columns[column].width) {
      deltaWidth = -1 * this.columns[column].width;
    }
    this.columns[column].width += deltaWidth;
    this.columns.forEach((col, index) => {
      if (index > column) {
        col.x += deltaWidth;
      }
    });
    this.refreshView();
  }

  /** 重新设置行高度 */
  resizeRow(row: number, deltaHeight: number) {
    if (!this.rows[row].height && deltaHeight < 0) {
      return;
    }
    if (deltaHeight < -1 * this.rows[row].height) {
      deltaHeight = -1 * this.rows[row].height;
    }
    this.rows[row].height += deltaHeight;
    this.rows.forEach((r, i) => {
      if (i > row) {
        r.y += deltaHeight;
      }
    });
    this.refreshView();
  }

  /** 自动滚动 */
  autoScroll(x: number, y: number) {
    if (
      x === this.mousePoint.x &&
      y === this.mousePoint.y &&
      this.state.isSelectCell
    ) {
      if (y > this.offsetTop + this.clientHeight || y < this.offsetTop) {
        this.scrollY(
          y < this.offsetTop
            ? -1 * this.style.cellHeight
            : this.style.cellHeight
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
          x < this.offsetLeft ? -1 * this.style.cellWidth : this.style.cellWidth
        );
        const range = this.calcActive(this.mousePoint.x, this.mousePoint.y);
        if (range) {
          this.activeArr.push(range);
          this.activeArr.splice(this.activeArr.length - 2, 1);
          this.setActive();
        }
      }
      this.autoScrollTimeoutID = setTimeout(() => {
        clearTimeout(this.autoScrollTimeoutID);
        if (this.mousePoint && this.state.isSelectCell) {
          this.autoScroll(this.mousePoint.x, this.mousePoint.y);
        }
      }, 50);
    } else {
      if (this.autoScrollTimeoutID) {
        clearTimeout(this.autoScrollTimeoutID);
        this.autoScrollTimeoutID = null;
      }
      this.autoScrollTimeoutID = setTimeout(() => {
        clearTimeout(this.autoScrollTimeoutID);
        if (this.mousePoint && this.state.isSelectCell) {
          this.autoScroll(this.mousePoint.x, this.mousePoint.y);
        }
      }, 50);
    }
  }

  /** 获取点击列标尺时选中范围 */
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
  }

  /** 获取点击列标尺时取消选中范围 */
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

  /** 获取点击行标尺时选中范围 */
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
  }

  /** 获取点击行标尺时取消选中范围 */
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
    if (scrollXThumbHeight < this.style.scrollBarThumbMinSize) {
      scrollXThumbHeight = this.style.scrollBarThumbMinSize;
    }
    const deltaX =
      ((x - this.mousePoint.x) * (this.scrollWidth - this.clientWidth)) /
      (this.clientWidth - scrollXThumbHeight);
    this.scrollX(deltaX);
  }

  calcScrollY(x: number, y: number) {
    let scrollYThumbHeight =
      (this.clientHeight / this.scrollHeight) * this.clientHeight;
    if (scrollYThumbHeight < this.style.scrollBarThumbMinSize) {
      scrollYThumbHeight = this.style.scrollBarThumbMinSize;
    }
    const deltaY =
      ((y - this.mousePoint.y) * (this.scrollHeight - this.clientHeight)) /
      (this.clientHeight - scrollYThumbHeight);
    this.scrollY(deltaY);
  }

  calcActive(x: number, y: number) {
    // for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
    //   for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
    //     const cell = this.viewCells[i][j].isCombined
    //       ? this.viewCells[i][j].combineCell
    //       : this.viewCells[i][j];
    //     if (this.inCellArea(x, y, cell)) {
    const cell = this.getViewCellByPoint(x, y);
    if (!cell) {
      return;
    }
    const range = {
      rowStart: this.activeCellPos.row,
      rowEnd: cell.position.row,
      columnStart: this.activeCellPos.column,
      columnEnd: cell.position.column,
    };
    this.recalcRange(range);
    return range;
    //     }
    //   }
    // }
  }

  calcUnActive(x: number, y: number) {
    // for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
    //   for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
    //     const cell = this.viewCells[i][j].isCombined
    //       ? this.viewCells[i][j].combineCell
    //       : this.viewCells[i][j];
    //     if (this.inCellArea(x, y, cell)) {
    const cell = this.getViewCellByPoint(x, y);
    if (!cell) {
      return;
    }
    const range = {
      rowStart: this.unActiveCellPos.row,
      rowEnd: cell.position.row,
      columnStart: this.unActiveCellPos.column,
      columnEnd: cell.position.column,
    };
    this.recalcRange(range);
    return range;
    //     }
    //   }
    // }
  }

  reCalcActiveArr() {
    const [rStart, rEnd, cStart, cEnd] = [
      Math.min(this.unActiveRange.rowStart, this.unActiveRange.rowEnd),
      Math.max(this.unActiveRange.rowStart, this.unActiveRange.rowEnd),
      Math.min(this.unActiveRange.columnStart, this.unActiveRange.columnEnd),
      Math.max(this.unActiveRange.columnStart, this.unActiveRange.columnEnd),
    ];

    let [rowStart, rowEnd, columnStart, columnEnd] = [
      rStart,
      rEnd,
      cStart,
      cEnd,
    ];
    for (
      let i = rStart, rLen = Math.min(rEnd, this.rows.length - 1);
      i <= rLen;
      i++
    ) {
      for (
        let j = cStart, cLen = Math.min(cEnd, this.columns.length - 1);
        j <= cLen;
        j++
      ) {
        const cell = this.cells[i][j].isCombined
          ? this.cells[i][j].combineCell
          : this.cells[i][j];
        if (cell.rowSpan > 1 || cell.colSpan > 1) {
          if (cell.position.row < rStart) {
            rowStart = cell.position.row;
          }
          if (cell.position.row + cell.rowSpan - 1 > rEnd) {
            rowEnd = cell.position.row + cell.rowSpan - 1;
          }
          if (cell.position.column < cStart) {
            columnStart = cell.position.column;
          }
          if (cell.position.column + cell.colSpan - 1 > cEnd) {
            columnEnd = cell.position.column + cell.colSpan - 1;
          }
        }
      }
    }

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
          rowEnd:
            this.cells[this.unActiveRange.rowStart][
              this.unActiveRange.columnStart
            ].position.row +
            this.cells[this.unActiveRange.rowStart][
              this.unActiveRange.columnStart
            ].rowSpan -
            1,
          columnStart: this.unActiveRange.columnStart,
          columnEnd:
            this.cells[this.unActiveRange.rowStart][
              this.unActiveRange.columnStart
            ].position.column +
            this.cells[this.unActiveRange.rowStart][
              this.unActiveRange.columnStart
            ].colSpan -
            1,
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
    let [row, column] = [Infinity, Infinity];
    this.activeArr.forEach((range) => {
      const rs = Math.min(range.rowStart, range.rowEnd);
      const cs = Math.min(range.columnStart, range.columnEnd);
      if (row > rs) {
        row = rs;
        column = cs;
      } else if (row === rs) {
        column = cs < column ? cs : column;
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

  onMouseUp = (event: MouseEvent) => {
    // event.preventDefault();
    // console.log('onmouseup');
    this.state.isSelectCell = false;
    this.state.isSelectScrollYThumb = false;
    this.state.isSelectScrollXThumb = false;
    this.state.isSelectRulerX = false;
    this.state.isSelectRulerY = false;
    this.state.isResizeColumn = false;
    this.state.isResizeRow = false;
    this.state.isMoveFloat = false;
    this.resizeColumnCell = null;
    this.resizeRowCell = null;
    this.state.isResizeFloat = null;
    this.state.resizeFloatPos = null;
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
    // console.log(this.activeArr);
    // console.log(this.activeCell);
  };

  createRow(count: number = 1) {
    // const lastModifyTime = new Date('1970-01-01').getTime();
    for (let i = 0; i < count; i++) {
      this.rows.push({
        y: this.rows.length
          ? this.rows[this.rows.length - 1].y +
            this.rows[this.rows.length - 1].height
          : 0,
        height: this.rows.length ? this.style.cellHeight : this.offsetTop,
        // fontWeight: {
        //   value: this.columns[i].style.fontWeight.value || this.style.cellFontWeight,
        //   lastModifyTime,
        // },
        // textAlign: {
        //   value:
        //     this.columns[i].style.textAlign.value ||
        //     (this.style.cellTextAlignLeft as CanvasTextAlign),
        //   lastModifyTime,
        // },
        // textBaseline: {
        //   value: this.style.cellTextBaseline as CanvasTextBaseline,
        //   lastModifyTime,
        // },
        // fontStyle: { value: this.style.cellFontStyle, lastModifyTime },
        // fontFamily: { value: this.style.cellFontFamily, lastModifyTime },
        // fontSize: { value: this.style.cellFontSize, lastModifyTime },
        // background: { value: this.style.cellBackgroundColor, lastModifyTime },
        // color: { value: this.style.cellColor, lastModifyTime },
        // borderWidth: { value: this.style.cellBorderWidth, lastModifyTime },
        // borderColor: { value: this.style.cellBorderColor, lastModifyTime },
      });
      this.cells.push(
        Array.from({ length: this.columns.length }).map((cv, ck) =>
          this.createCell(this.cells.length, ck)
        )
      );
    }
  }

  createColumn(count: number) {
    // const lastModifyTime = new Date().getTime();
    for (let i = 0; i < count; i++) {
      const columnLength = this.columns.length;
      this.columns.push({
        x: columnLength
          ? this.columns[columnLength - 1].x +
            this.columns[columnLength - 1].width
          : 0,
        width: columnLength ? this.style.cellWidth : this.offsetLeft,
        // fontWeight: { value: this.style.cellFontWeight, lastModifyTime },
        // textAlign: {
        //   value: this.style.cellTextAlignLeft as CanvasTextAlign,
        //   lastModifyTime,
        // },
        // textBaseline: {
        //   value: this.style.cellTextBaseline as CanvasTextBaseline,
        //   lastModifyTime,
        // },
        // fontStyle: { value: this.style.cellFontStyle, lastModifyTime },
        // fontFamily: { value: this.style.cellFontFamily, lastModifyTime },
        // fontSize: { value: this.style.cellFontSize, lastModifyTime },
        // background: { value: this.style.cellBackgroundColor, lastModifyTime },
        // color: { value: this.style.cellColor, lastModifyTime },
        // borderWidth: { value: this.style.cellBorderWidth, lastModifyTime },
        // borderColor: { value: this.style.cellBorderColor, lastModifyTime },
      });
      this.cells.forEach((row, rk) =>
        row.push(this.createCell(rk, columnLength))
      );
    }
  }

  scrollX(deltaX: number, immediate = true) {
    if (this.scrollLeft === 0 && deltaX < 0) {
      return;
    }
    this.scrollLeft += deltaX;
    if (this.scrollLeft >= this.scrollWidth - this.clientWidth) {
      this.createColumn(
        Math.ceil(
          (this.scrollLeft - this.scrollWidth + this.clientWidth) /
            this.style.cellWidth
        ) + 1
      );
      this.scrollWidth =
        this.cells[0][this.cells[0].length - 1].x +
        this.cells[0][this.cells[0].length - 1].width -
        this.offsetLeft;
      // this.scrollLeft = this.scrollWidth - this.clientWidth - this.style.cellWidth;
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
      this.createRow(
        Math.ceil(
          (this.scrollTop - this.scrollHeight + this.clientHeight) /
            this.style.cellHeight
        ) + 1
      );
      this.scrollHeight =
        this.cells[this.cells.length - 1][0].y +
        this.cells[this.cells.length - 1][0].height -
        this.offsetTop;
      // this.scrollTop = this.scrollHeight - this.clientHeight - this.style.cellHeight;
    } else if (this.scrollTop <= 0) {
      this.scrollTop = 0;
    }
    if (immediate) {
      this.refreshView();
    }
  }

  onWheel(event: WheelEvent) {
    // console.log('wheel', event);
    // if (!this.isTicking) {
    //   requestAnimationFrame(() => {
    this.scrollY(event.deltaY);

    //     this.isTicking = false;
    //   });
    // }
    // this.isTicking = true;
  }

  onKeyArrowUpOrDown(event: KeyboardEvent) {
    if (!this.activeArr || !this.activeArr.length) {
      const floatArr = this.floatArr.filter((elem) => elem.isActive);
      for (const elem of floatArr) {
        elem.y += event.code === KeyboardCode.ArrowDown ? 10 : -10;
        elem.y = elem.y < this.offsetTop ? this.offsetTop : elem.y;
      }
      this.refreshView();
      return;
    }
    if (!event.shiftKey) {
      if (event.code === KeyboardCode.ArrowDown) {
        const step =
          this.activeCell.position.row +
          this.activeCell.rowSpan -
          this.activeCellPos.row;
        this.activeCellPos = {
          row:
            this.activeCellPos.row + step > this.cells.length - 1 ||
            event.ctrlKey
              ? this.cells.length - 1
              : this.activeCellPos.row + step,
          column: this.activeCellPos.column,
          rangeIndex: 0,
        };
      } else if (event.code === KeyboardCode.ArrowUp) {
        const step = this.activeCellPos.row - this.activeCell.position.row + 1;
        this.activeCellPos = {
          row:
            this.activeCellPos.row - step < 1 || event.ctrlKey
              ? 1
              : this.activeCellPos.row - step,
          column: this.activeCellPos.column,
          rangeIndex: 0,
        };
      }
      this.activeArr = [
        {
          rowStart: this.activeCell.position.row,
          rowEnd: this.activeCell.position.row + this.activeCell.rowSpan - 1,
          columnStart: this.activeCell.position.column,
          columnEnd:
            this.activeCell.position.column + this.activeCell.colSpan - 1,
        },
      ];
      // if (!this.isTicking) {
      //   requestAnimationFrame(() => {
      this.resetCellPerspective(
        this.cells[this.activeCellPos.row][this.activeCellPos.column]
      );

      //     this.isTicking = false;
      //   });
      // }
    } else {
      const range = this.activeArr[this.activeCellPos.rangeIndex];
      if (range && range.rowEnd !== Infinity) {
        if (
          this.activeCellPos.row === range.rowEnd &&
          range.rowStart !== range.rowEnd
        ) {
          const temp = range.rowStart;
          range.rowStart = range.rowEnd;
          range.rowEnd = temp;
        }
        if (event.code === KeyboardCode.ArrowDown) {
          const isExpand = range.rowEnd >= range.rowStart;
          range.rowEnd =
            range.rowEnd + 1 > this.rows.length - 1 || event.ctrlKey
              ? this.rows.length - 1
              : range.rowEnd + 1;
          this.recalcRange(range, isExpand);
        } else if (event.code === KeyboardCode.ArrowUp) {
          const isExpand = range.rowEnd <= range.rowStart;
          range.rowEnd =
            range.rowEnd - 1 < 1 || event.ctrlKey ? 1 : range.rowEnd - 1;
          this.recalcRange(range, isExpand);
        }
        // if (!this.isTicking) {
        //   requestAnimationFrame(() => {
        this.resetCellPerspective(
          this.cells[range.rowEnd][
            range.columnEnd === Infinity ? 0 : range.columnEnd
          ]
        );

        //       this.isTicking = false;
        //     });
        //   }
        //   this.isTicking = true;
      }
    }
  }

  onKeyArrowLeftOrRight(event: KeyboardEvent) {
    if (!this.activeArr || !this.activeArr.length) {
      const floatArr = this.floatArr.filter((elem) => elem.isActive);
      for (const elem of floatArr) {
        elem.x += event.code === KeyboardCode.ArrowRight ? 10 : -10;
        elem.x = elem.x < this.offsetLeft ? this.offsetLeft : elem.x;
      }
      this.refreshView();
      return;
    }
    if (!event.shiftKey) {
      if (event.code === KeyboardCode.ArrowRight) {
        const step =
          this.activeCell.position.column +
          this.activeCell.colSpan -
          this.activeCellPos.column;
        this.activeCellPos = {
          row: this.activeCellPos.row,
          column:
            this.activeCellPos.column + step > this.cells[0].length - 1 ||
            event.ctrlKey
              ? this.cells[0].length - 1
              : this.activeCellPos.column + step,
          rangeIndex: 0,
        };
      } else if (event.code === KeyboardCode.ArrowLeft) {
        const step =
          this.activeCellPos.column - this.activeCell.position.column + 1;
        this.activeCellPos = {
          row: this.activeCellPos.row,
          column:
            this.activeCellPos.column - step < 1 || event.ctrlKey
              ? 1
              : this.activeCellPos.column - step,
          rangeIndex: 0,
        };
      }
      this.activeArr = [
        {
          rowStart: this.activeCell.position.row,
          rowEnd: this.activeCell.position.row + this.activeCell.rowSpan - 1,
          columnStart: this.activeCell.position.column,
          columnEnd:
            this.activeCell.position.column + this.activeCell.colSpan - 1,
        },
      ];
      // if (!this.isTicking) {
      //   requestAnimationFrame(() => {
      this.resetCellPerspective(
        this.cells[this.activeCellPos.row][this.activeCellPos.column]
      );
      //   this.isTicking = false;
      // });
      // }
    } else {
      const range = this.activeArr[this.activeCellPos.rangeIndex];
      if (range && range.columnEnd !== Infinity) {
        if (
          this.activeCellPos.column === range.columnEnd &&
          range.columnStart !== range.columnEnd
        ) {
          const temp = range.columnStart;
          range.columnStart = range.columnEnd;
          range.columnEnd = temp;
        }
        if (event.code === KeyboardCode.ArrowRight) {
          const isExpand = range.columnEnd >= range.columnStart;
          range.columnEnd =
            range.columnEnd + 1 > this.cells[0].length - 1 || event.ctrlKey
              ? this.cells[0].length - 1
              : range.columnEnd + 1;
          this.recalcRange(range, isExpand);
        } else if (event.code === KeyboardCode.ArrowLeft) {
          const isExpand = range.columnEnd <= range.columnStart;
          range.columnEnd =
            range.columnEnd - 1 < 1 || event.ctrlKey ? 1 : range.columnEnd - 1;
          this.recalcRange(range, isExpand);
        }
        // if (!this.isTicking) {
        //   requestAnimationFrame(() => {
        this.resetCellPerspective(
          this.cells[range.rowEnd === Infinity ? 0 : range.rowEnd][
            event.code === KeyboardCode.ArrowRight
              ? Math.max(range.columnStart, range.columnEnd)
              : Math.min(range.columnStart, range.columnEnd)
          ]
        );
        //     this.isTicking = false;
        //   });
        // }
        // this.isTicking = true;
      }
    }
  }

  onKeyTabOrEnter(event: KeyboardEvent) {
    if (
      this.activeArr.length === 1 &&
      this.activeCell.position.row ===
        Math.min(this.activeArr[0].rowStart, this.activeArr[0].rowEnd) &&
      this.activeCell.position.column ===
        Math.min(this.activeArr[0].columnStart, this.activeArr[0].columnEnd) &&
      this.activeCell.position.row + this.activeCell.rowSpan - 1 ===
        Math.max(this.activeArr[0].rowStart, this.activeArr[0].rowEnd) &&
      this.activeCell.position.column + this.activeCell.colSpan - 1 ===
        Math.max(this.activeArr[0].columnStart, this.activeArr[0].columnEnd)
    ) {
      if (!event.shiftKey) {
        if (event.code === KeyboardCode.Tab) {
          const step =
            this.activeCell.position.column +
            this.activeCell.colSpan -
            this.activeCellPos.column;
          this.activeCellPos = {
            row: this.activeCellPos.row,
            column:
              this.activeCellPos.column + step > this.cells[0].length - 1 ||
              event.ctrlKey
                ? this.cells[0].length - 1
                : this.activeCellPos.column + step,
            rangeIndex: 0,
          };
        } else if (event.code === KeyboardCode.Enter) {
          const step =
            this.activeCell.position.row +
            this.activeCell.rowSpan -
            this.activeCellPos.row;
          this.activeCellPos = {
            row:
              this.activeCellPos.row + step > this.cells.length - 1 ||
              event.ctrlKey
                ? this.cells.length - 1
                : this.activeCellPos.row + step,
            column: this.activeCellPos.column,
            rangeIndex: 0,
          };
        }
      } else {
        if (event.code === KeyboardCode.Tab) {
          const step =
            this.activeCellPos.column - this.activeCell.position.column + 1;
          this.activeCellPos = {
            row: this.activeCellPos.row,
            column:
              this.activeCellPos.column - 1 < 1 || event.ctrlKey
                ? 1
                : this.activeCellPos.column - 1,
            rangeIndex: 0,
          };
        } else if (event.code === KeyboardCode.Enter) {
          const step =
            this.activeCellPos.row - this.activeCell.position.row + 1;
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
          rowStart: this.activeCell.position.row,
          rowEnd: this.activeCell.position.row + this.activeCell.rowSpan - 1,
          columnStart: this.activeCell.position.column,
          columnEnd:
            this.activeCell.position.column + this.activeCell.colSpan - 1,
        },
      ];
      // if (!this.isTicking) {
      //   requestAnimationFrame(() => {
      this.resetCellPerspective(
        this.cells[this.activeCellPos.row][this.activeCellPos.column]
      );
      //     this.isTicking = false;
      //   });
      // }
    } else {
      let { row, column, rangeIndex } = this.activeCellPos;
      let next = false;
      if (event.code === KeyboardCode.Tab) {
        column = event.shiftKey
          ? this.activeCell.position.column - 1
          : this.activeCell.position.column + this.activeCell.colSpan;
        while (true) {
          const range = this.activeArr[rangeIndex];
          const rowStart = Math.min(range.rowStart, range.rowEnd);
          const rowEnd = Math.max(range.rowStart, range.rowEnd);
          const columnStart = Math.min(range.columnStart, range.columnEnd);
          const columnEnd = Math.max(range.columnStart, range.columnEnd);
          if (!event.shiftKey) {
            do {
              if (
                this.cells[row][column].position.column +
                  this.cells[row][column].colSpan -
                  1 >
                columnEnd
              ) {
                row++;
                if (row > rowEnd) {
                  break;
                }
                column = columnStart;
              } else if (!this.cells[row][column].isCombined) {
                this.activeCellPos = { row, column, rangeIndex };
                next = true;
                break;
              } else {
                column += this.cells[row][column].colSpan;
              }
            } while (true);
            if (!next) {
              rangeIndex =
                rangeIndex + 1 > this.activeArr.length - 1 ? 0 : rangeIndex + 1;
              row = Math.min(
                this.activeArr[rangeIndex].rowStart,
                this.activeArr[rangeIndex].rowEnd
              );
              column = Math.min(
                this.activeArr[rangeIndex].columnStart,
                this.activeArr[rangeIndex].columnEnd
              );
            } else {
              break;
            }
          } else {
            do {
              if (this.cells[row][column].position.column < columnStart) {
                row--;
                if (row < rowStart) {
                  break;
                }
                column = columnEnd;
              } else if (!this.cells[row][column].isCombined) {
                this.activeCellPos = { row, column, rangeIndex };
                next = true;
                break;
              } else {
                column--;
              }
            } while (true);
            if (!next) {
              rangeIndex =
                rangeIndex - 1 < 0 ? this.activeArr.length - 1 : rangeIndex - 1;
              row = Math.max(
                this.activeArr[rangeIndex].rowStart,
                this.activeArr[rangeIndex].rowEnd
              );
              column = Math.max(
                this.activeArr[rangeIndex].columnStart,
                this.activeArr[rangeIndex].columnEnd
              );
            } else {
              break;
            }
          }
        }
      } else if (event.code === KeyboardCode.Enter) {
        row = event.shiftKey
          ? this.activeCell.position.row - 1
          : this.activeCell.position.row + this.activeCell.rowSpan;
        while (true) {
          const range = this.activeArr[rangeIndex];
          const rowStart = Math.min(range.rowStart, range.rowEnd);
          const rowEnd = Math.max(range.rowStart, range.rowEnd);
          const columnStart = Math.min(range.columnStart, range.columnEnd);
          const columnEnd = Math.max(range.columnStart, range.columnEnd);
          if (!event.shiftKey) {
            do {
              if (
                this.cells[row][column].position.row +
                  this.cells[row][column].rowSpan -
                  1 >
                rowEnd
              ) {
                column++;
                if (column > columnEnd) {
                  break;
                }
                row = rowStart;
              } else if (!this.cells[row][column].isCombined) {
                this.activeCellPos = { row, column, rangeIndex };
                next = true;
                break;
              } else {
                row++;
              }
            } while (true);
            if (!next) {
              rangeIndex =
                rangeIndex + 1 > this.activeArr.length - 1 ? 0 : rangeIndex + 1;
              row = Math.min(
                this.activeArr[rangeIndex].rowStart,
                this.activeArr[rangeIndex].rowEnd
              );
              column = Math.min(
                this.activeArr[rangeIndex].columnStart,
                this.activeArr[rangeIndex].columnEnd
              );
            } else {
              break;
            }
          } else {
            do {
              if (this.cells[row][column].position.row < rowStart) {
                column--;
                if (column < columnStart) {
                  break;
                }
                row = rowEnd;
              } else if (!this.cells[row][column].isCombined) {
                this.activeCellPos = { row, column, rangeIndex };
                next = true;
                break;
              } else {
                row--;
              }
            } while (true);
            if (!next) {
              rangeIndex =
                rangeIndex - 1 < 0 ? this.activeArr.length - 1 : rangeIndex - 1;
              row = Math.max(
                this.activeArr[rangeIndex].rowStart,
                this.activeArr[rangeIndex].rowEnd
              );
              column = Math.max(
                this.activeArr[rangeIndex].columnStart,
                this.activeArr[rangeIndex].columnEnd
              );
            } else {
              break;
            }
          }
        }
      }

      // if (!this.isTicking) {
      //   requestAnimationFrame(() => {
      this.resetCellPerspective(
        this.cells[this.activeCellPos.row][this.activeCellPos.column]
      );
      //     this.isTicking = false;
      //   });
      // }
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
            ? this.activeCellPos.row +
              this.cells[this.activeCellPos.row][this.activeCellPos.column]
                .rowSpan -
              1
            : this.cells[1][1].position.row + this.cells[1][1].rowSpan - 1,
        columnStart: event.shiftKey ? this.activeCellPos.column : 1,
        columnEnd:
          event.shiftKey || event.ctrlKey || (!event.shiftKey && !event.ctrlKey)
            ? this.cells[1][1].position.column + this.cells[1][1].colSpan - 1
            : this.activeCellPos.column +
              this.cells[this.activeCellPos.row][this.activeCellPos.column]
                .colSpan -
              1,
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
      (event.code === KeyboardCode.PageDown ? 1 : -1) * this.clientHeight;
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
              (i - this.cells.length) * this.style.cellHeight,
            this.style.cellHeight,
          ];
      if (inRange(posY, y, y + height, true)) {
        const cell = this.cells[i][this.activeCellPos.column].isCombined
          ? this.cells[i][this.activeCellPos.column].combineCell
          : this.cells[i][this.activeCellPos.column];
        this.scrollY(y - this.cells[this.activeCellPos.row][0].y, false);
        this.activeArr = [
          {
            rowStart: event.shiftKey
              ? this.activeCellPos.row
              : cell.position.row,
            rowEnd: cell.position.row + cell.rowSpan - 1,
            columnStart: Math.min(
              this.activeArr[this.activeCellPos.rangeIndex].columnStart,
              cell.position.column
            ),
            columnEnd: event.shiftKey
              ? Math.max(
                  this.activeArr[this.activeCellPos.rangeIndex].columnEnd,
                  cell.position.column + cell.colSpan - 1
                )
              : cell.position.column + cell.colSpan - 1,
          },
        ];
        this.activeCellPos = {
          row: event.shiftKey ? this.activeCellPos.row : i,
          column: this.activeCellPos.column,
          rangeIndex: 0,
        };
        // if (!this.isTicking) {
        //   requestAnimationFrame(() => {
        this.resetCellPerspective(
          event.shiftKey
            ? this.cells[
                event.code === KeyboardCode.PageDown
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
        //     this.isTicking = false;
        //   });
        // }
        // this.isTicking = true;
        break;
      } else {
        event.code === KeyboardCode.PageDown ? i++ : i--;
      }
    }
  }

  deleteContent(rangeArr: CellRange[]) {
    this.clipBoard = null;
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
    this.drawRuler(this.ctx);
    this.drawScrollBar(this.ctx);
  }

  onKeyDown(event: KeyboardEvent) {
    // console.log('keydown', event);
    switch (event.code) {
      case KeyboardCode.Tab:
      case KeyboardCode.Enter:
        event.preventDefault();
        this.onKeyTabOrEnter(event);
        this.refreshView();
        break;
      case KeyboardCode.ArrowUp:
      case KeyboardCode.ArrowDown:
        event.preventDefault();
        this.onKeyArrowUpOrDown(event);
        break;
      case KeyboardCode.ArrowLeft:
      case KeyboardCode.ArrowRight:
        event.preventDefault();
        this.onKeyArrowLeftOrRight(event);
        break;
      case KeyboardCode.Home:
        event.preventDefault();
        this.onKeyHome(event);
        break;
      case KeyboardCode.PageUp:
      case KeyboardCode.PageDown:
        event.preventDefault();
        this.onKeyPageUpOrDown(event);
        break;
      case KeyboardCode.Delete:
        if (this.floatArr.find((elem) => elem.isActive)) {
          for (let i = this.floatArr.length - 1; i >= 0; i--) {
            if (this.floatArr[i].isActive) {
              this.floatArr.splice(i, 1);
            }
          }
          this.activeCellPos = { row: 1, column: 1, rangeIndex: 0 };
          this.activeArr = [
            { rowStart: 1, columnStart: 1, rowEnd: 1, columnEnd: 1 },
          ];
          this.drawFloat();
          this.setActive();
        } else {
          this.deleteContent(this.activeArr);
        }
        break;
      case KeyboardCode.Backspace:
        this.resetCellPerspective(
          this.cells[this.activeCellPos.row][this.activeCellPos.column]
        );
        this.clipBoard = null;
        this.editingCell = this.cells[this.activeCellPos.row][
          this.activeCellPos.column
        ];
        this.editingCell.content.value = null;
        this.state.isCellEdit = true;
        break;
      case KeyboardCode.KeyA:
        if (event.ctrlKey) {
          event.preventDefault();
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
        }
        break;
      case KeyboardCode.KeyC:
        if (event.ctrlKey) {
          event.preventDefault();
          if (this.activeArr.length === 1) {
            this.clipBoard = {
              range: cloneDeep(this.activeArr[0]),
              isCut: false,
            };
            this.setClipStatusAnimation();
          } else if (this.activeArr.length) {
            alert('cannot copy');
          }
        }
        break;
      case KeyboardCode.KeyX:
        if (event.ctrlKey) {
          event.preventDefault();
          if (this.activeArr.length === 1) {
            this.clipBoard = {
              range: cloneDeep(this.activeArr[0]),
              isCut: true,
            };
            this.setClipStatusAnimation();
          } else if (this.activeArr.length) {
            alert('cannot cut');
          }
        }
        break;
      case KeyboardCode.KeyV:
        if (event.ctrlKey) {
          // event.preventDefault();
          if (!this.paste()) {
            event.preventDefault();
          }
        }
        break;
      case KeyboardCode.Escape:
        this.clipBoard = null;
        break;
      default:
        break;
    }
  }

  onKeyPress(event: KeyboardEvent) {
    // console.log('keypress', event);
    event.preventDefault();
    if (!event.ctrlKey) {
      this.resetCellPerspective(
        this.cells[this.activeCellPos.row][this.activeCellPos.column]
      );
      this.clipBoard = null;
      this.editingCell = this.cells[this.activeCellPos.row][
        this.activeCellPos.column
      ];
      this.editingCell.content.value = event.key;
      this.editingCell.content.html = event.key;
      this.state.isCellEdit = true;
    }
  }

  paste() {
    if (!this.clipBoard) {
      return true;
    }
    if (this.clipBoard.isCut && this.activeArr.length > 1) {
      alert('cannot paste');
      return false;
    }

    const [clipRowStart, clipRowEnd, clipColumnStart, clipColumnEnd] = [
      Math.min(this.clipBoard.range.rowStart, this.clipBoard.range.rowEnd),
      Math.min(
        this.rows.length - 1,
        Math.max(this.clipBoard.range.rowStart, this.clipBoard.range.rowEnd)
      ),
      Math.min(
        this.clipBoard.range.columnStart,
        this.clipBoard.range.columnEnd
      ),
      Math.min(
        this.columns.length - 1,
        Math.max(
          this.clipBoard.range.columnStart,
          this.clipBoard.range.columnEnd
        )
      ),
    ];
    const clipCells = cloneDeep(
      this.cells
        .filter((row, rowIndex) =>
          inRange(rowIndex, clipRowStart, clipRowEnd, true)
        )
        .map((row) =>
          row.filter((cell, columnIndex) =>
            inRange(columnIndex, clipColumnStart, clipColumnEnd, true)
          )
        )
    );

    if (!(clipRowStart === clipRowEnd && clipColumnStart === clipColumnEnd)) {
      for (let m = 0, len = this.activeArr.length; m < len; m++) {
        const range = this.activeArr[m];
        const rowStart = Math.min(range.rowStart, range.rowEnd);
        const rowEnd = rowStart + clipRowEnd - clipRowStart;
        if (rowEnd > this.rows.length - 1) {
          this.createRow(rowEnd - this.rows.length + 1);
        }
        const columnStart = Math.min(range.columnStart, range.columnEnd);
        const columnEnd = columnStart + clipColumnEnd - clipColumnStart;
        if (columnEnd > this.columns.length - 1) {
          this.createColumn(columnEnd - this.columns.length + 1);
        }
        let breakCombine = false;

        for (let i = rowStart; i <= rowEnd; i++) {
          if (
            (this.cells[i][columnStart].isCombined &&
              this.cells[i][columnStart].combineCell.position.column <
                columnStart) ||
            (
              (this.cells[i][columnEnd].isCombined &&
                this.cells[i][columnEnd].combineCell) ||
              this.cells[i][columnEnd]
            ).position.column +
              (
                (this.cells[i][columnEnd].isCombined &&
                  this.cells[i][columnEnd].combineCell) ||
                this.cells[i][columnEnd]
              ).colSpan -
              1 >
              columnEnd
          ) {
            breakCombine = true;
            break;
          }
        }
        if (!breakCombine) {
          for (let i = columnStart; i <= columnEnd; i++) {
            if (
              (this.cells[rowStart][i].isCombined &&
                this.cells[rowStart][i].combineCell.position.row < rowStart) ||
              (
                (this.cells[rowEnd][i].isCombined &&
                  this.cells[rowEnd][i].combineCell) ||
                this.cells[rowEnd][i]
              ).position.row +
                (
                  (this.cells[rowEnd][i].isCombined &&
                    this.cells[rowEnd][i].combineCell) ||
                  this.cells[rowEnd][i]
                ).rowSpan -
                1 >
                rowEnd
            ) {
              breakCombine = true;
              break;
            }
          }
        }

        if (breakCombine && !this.clipBoard.isCut) {
          alert('cannot paste');
          return false;
        } else if (
          breakCombine &&
          this.clipBoard.isCut &&
          !confirm('operation will split combinded cells')
        ) {
          return false;
        }
      }
    }
    for (let m = 0, len = this.activeArr.length; m < len; m++) {
      const range = this.activeArr[m];
      const rowStart = Math.min(range.rowStart, range.rowEnd);
      const rowEnd =
        clipRowStart === clipRowEnd && clipColumnStart === clipColumnEnd
          ? Math.max(range.rowStart, range.rowEnd)
          : rowStart + clipRowEnd - clipRowStart;
      if (rowEnd > this.rows.length - 1) {
        this.createRow(rowEnd - this.rows.length + 1);
      }
      const columnStart = Math.min(range.columnStart, range.columnEnd);
      const columnEnd =
        clipRowStart === clipRowEnd && clipColumnStart === clipColumnEnd
          ? Math.max(range.columnStart, range.columnEnd)
          : columnStart + clipColumnEnd - clipColumnStart;
      if (columnEnd > this.columns.length - 1) {
        this.createColumn(columnEnd - this.columns.length + 1);
      }
      for (let i = rowStart; i <= rowEnd; i++) {
        for (let j = columnStart; j <= columnEnd; j++) {
          const clipCell =
            clipRowStart === clipRowEnd && clipColumnStart === clipColumnEnd
              ? clipCells[0][0]
              : clipCells[i - rowStart][j - columnStart];
          if (
            this.cells[i][j].isCombined ||
            this.cells[i][j].rowSpan > 1 ||
            this.cells[i][j].colSpan > 1
          ) {
            this.unCombine(
              (this.cells[i][j].isCombined && this.cells[i][j].combineCell) ||
                this.cells[i][j]
            );
          }
          this.cells[i][j].style = clipCell.style;
          this.cells[i][j].content = clipCell.content;
          // this.cells[i][j].isCombined = clipCell.isCombined;
          this.cells[i][j].rowSpan = clipCell.rowSpan;
          this.cells[i][j].colSpan = clipCell.colSpan;
          this.cells[i][j].combineCell =
            (clipCell.combineCell &&
              this.cells[rowStart + clipCell.combineCell.position.row - 1][
                rowEnd + clipCell.combineCell.position.column
              ]) ||
            null;
        }
      }

      if (this.clipBoard.isCut) {
        for (let i = clipRowStart; i <= clipRowEnd; i++) {
          for (let j = clipColumnStart; j <= clipColumnEnd; j++) {
            if (
              inRange(i, rowStart, rowEnd, true) &&
              inRange(j, columnStart, columnEnd, true)
            ) {
              continue;
            }
            this.cells[i][j] = this.createCell(i, j);
          }
        }
        // this.clipBoard = null;
      } else if (
        (inRange(rowStart, clipRowStart, clipRowEnd, true) ||
          inRange(rowEnd, clipRowStart, clipRowEnd, true)) &&
        (inRange(columnStart, clipColumnStart, clipColumnEnd, true) ||
          inRange(columnEnd, clipColumnStart, clipColumnEnd))
      ) {
        // this.clipBoard = null;
      }

      this.activeArr[m] = { rowStart, rowEnd, columnStart, columnEnd };
    }
    if (
      this.activeArr.find(
        (range) =>
          (inRange(range.rowStart, clipRowStart, clipRowEnd, true) ||
            inRange(range.rowEnd, clipRowStart, clipRowEnd, true)) &&
          (inRange(range.columnStart, clipColumnStart, clipColumnEnd, true) ||
            inRange(range.columnEnd, clipColumnStart, clipColumnEnd, true))
      )
    ) {
      this.clipBoard = null;
    }

    this.resetCellPerspective(this.activeCell);
    return true;
  }

  editCellCompelte(change = true) {
    if (!change) {
      this.editingCell.content.value = this.editingCell.content.previousValue;
      this.editingCell.content.html = this.editingCell.content.previousHtml;
    } else if (this.editingCell) {
      this.editingCell.content.previousValue = this.editingCell.content.value;
      this.editingCell.content.previousHtml = this.editingCell.content.html;
      this.drawCell(this.ctx, this.editingCell, true);
      this.drawRuler(this.ctx);
      this.drawScrollBar(this.ctx);
    }
    this.editingCell = null;
    this.state.isCellEdit = false;
    this.refreshView();
    this.canvas.focus();
  }

  onEditCellKeyDown(event: KeyboardEvent) {
    switch (event.code) {
      case KeyboardCode.Tab:
      case KeyboardCode.Enter:
        this.editCellCompelte();
        event.preventDefault();
        this.onKeyTabOrEnter(event);
        break;
      case KeyboardCode.ArrowUp:
      case KeyboardCode.ArrowDown:
        if (!this.editingCell.content || !this.editingCell.content.value) {
          this.editCellCompelte();
          event.preventDefault();
          this.onKeyArrowUpOrDown(event);
        }
        break;
      case KeyboardCode.ArrowLeft:
      case KeyboardCode.ArrowRight:
        if (!this.editingCell.content || !this.editingCell.content.value) {
          this.editCellCompelte();
          event.preventDefault();
          this.onKeyArrowLeftOrRight(event);
        }
        break;
      case KeyboardCode.Escape:
        this.clipBoard = null;
        this.editCellCompelte(false);
        break;
      default:
        break;
    }
  }

  unCombine(combineCell: Cell) {
    const mLen = Math.min(
      combineCell.position.row + combineCell.rowSpan,
      this.rows.length
    );
    const nLen = Math.min(
      combineCell.position.column + combineCell.colSpan,
      this.columns.length
    );
    for (let m = combineCell.position.row, rowLen = mLen; m < rowLen; m++) {
      for (
        let n = combineCell.position.column, colLen = nLen;
        n < colLen;
        n++
      ) {
        // this.cells[m][n].isCombined = false;
        this.cells[m][n].rowSpan = 1;
        this.cells[m][n].colSpan = 1;
        this.cells[m][n].combineCell = null;
      }
    }
  }

  toggleCombine() {
    for (let index = 0, len = this.activeArr.length; index < len; index++) {
      const range = this.activeArr[index];
      if (
        range.rowStart !== range.rowEnd ||
        range.columnStart !== range.columnEnd
      ) {
        const rowStart = Math.min(range.rowStart, range.rowEnd);
        const rowEnd = Math.max(range.rowStart, range.rowEnd);
        const columnStart = Math.min(range.columnStart, range.columnEnd);
        const columnEnd = Math.max(range.columnStart, range.columnEnd);
        const cell = this.cells[rowStart][columnStart];
        let hasCombine = false;
        for (
          let i = rowStart, rLen = Math.min(this.rows.length - 1, rowEnd);
          i <= rLen;
          i++
        ) {
          for (
            let j = columnStart,
              cLen = Math.min(this.columns.length - 1, columnEnd);
            j <= cLen;
            j++
          ) {
            if (
              this.cells[i][j].rowSpan > 1 ||
              this.cells[i][j].colSpan > 1 ||
              this.cells[i][j].isCombined
            ) {
              hasCombine = true;

              const combineCell = this.cells[i][j].isCombined
                ? this.cells[i][j].combineCell
                : this.cells[i][j];
              this.unCombine(combineCell);
            }
          }
        }
        if (!hasCombine) {
          cell.rowSpan = rowEnd - rowStart + 1;
          cell.colSpan = columnEnd - columnStart + 1;
          for (
            let i = rowStart, rLen = Math.min(this.rows.length - 1, rowEnd);
            i <= rLen;
            i++
          ) {
            for (
              let j = columnStart,
                cLen = Math.min(this.columns.length - 1, columnEnd);
              j <= cLen;
              j++
            ) {
              if (i === rowStart && j === columnStart) {
                continue;
              }
              // this.cells[i][j].isCombined = true;
              this.cells[i][j].rowSpan = 1;
              this.cells[i][j].colSpan = 1;
              this.cells[i][j].combineCell = cell;
            }
          }
          this.activeCellPos = {
            row: cell.position.row,
            column: cell.position.column,
          };
        }
      } else if (
        this.cells[range.rowStart][range.columnStart].rowSpan > 1 ||
        this.cells[range.rowStart][range.columnStart].colSpan > 1
      ) {
        const cell = this.cells[range.rowStart][range.columnStart];
        const rowStart = range.rowStart;
        const rowEnd = range.rowStart + cell.rowSpan - 1;
        const columnStart = range.columnStart;
        const columnEnd = range.columnEnd + cell.colSpan - 1;
        for (let i = rowStart; i <= rowEnd; i++) {
          for (let j = columnStart; j <= columnEnd; j++) {
            // this.cells[i][j].isCombined = false;
            this.cells[i][j].rowSpan = 1;
            this.cells[i][j].colSpan = 1;
            this.cells[i][j].combineCell = null;
          }
        }
      }
      this.refreshView();
      this.canvas.focus();
    }
  }

  recalcRange(range: CellRange, isExpand = true) {
    const rowReverse = range.rowStart > range.rowEnd;
    const columnReverse = range.columnStart > range.columnEnd;
    let needReCalc = false;
    const [rStart, rEnd, cStart, cEnd] = [
      Math.min(range.rowStart, range.rowEnd),
      Math.max(range.rowStart, range.rowEnd),
      Math.min(range.columnStart, range.columnEnd),
      Math.max(range.columnStart, range.columnEnd),
    ];

    let { rowStart, rowEnd, columnStart, columnEnd } = range;
    for (
      let i = rStart, rLen = Math.min(rEnd, this.rows.length - 1);
      i <= rLen;
      i++
    ) {
      for (
        let j = cStart, cLen = Math.min(cEnd, this.columns.length - 1);
        j <= cLen;
        j++
      ) {
        if (this.cells[i][j].hidden) {
          continue;
        }
        const cell = this.cells[i][j].isCombined
          ? this.cells[i][j].combineCell
          : this.cells[i][j];
        if (cell.hidden || (cell.rowSpan === 1 && cell.colSpan === 1)) {
          continue;
        }
        if (cell.position.row < rStart) {
          if (isExpand) {
            rowReverse
              ? (rowEnd = cell.position.row)
              : (rowStart = cell.position.row);
          } else {
            if (cell.position.row + cell.rowSpan > rEnd) {
              rowStart = cell.position.row;
              rowEnd = cell.position.row + cell.rowSpan;
            } else {
              rowEnd = cell.position.row + cell.rowSpan;
            }
          }
          needReCalc = true;
        }
        if (cell.position.row + cell.rowSpan - 1 > rEnd) {
          if (isExpand) {
            rowReverse
              ? (rowStart = cell.position.row + cell.rowSpan - 1)
              : (rowEnd = cell.position.row + cell.rowSpan - 1);
          } else {
            if (cell.position.row - 1 < rStart) {
              rowStart = cell.position.row + cell.rowSpan - 1;
              rowEnd = cell.position.row - 1;
            } else {
              rowEnd = cell.position.row - 1;
            }
          }
          needReCalc = true;
        }
        if (cell.position.column < cStart) {
          if (isExpand) {
            columnReverse
              ? (columnEnd = cell.position.column)
              : (columnStart = cell.position.column);
          } else {
            if (cell.position.column + cell.colSpan > cEnd) {
              columnStart = cell.position.column;
              columnEnd = cell.position.column + cell.colSpan;
            } else {
              columnEnd = cell.position.column + cell.colSpan;
            }
          }
          needReCalc = true;
        }
        if (cell.position.column + cell.colSpan - 1 > cEnd) {
          if (isExpand) {
            columnReverse
              ? (columnStart = cell.position.column + cell.colSpan - 1)
              : (columnEnd = cell.position.column + cell.colSpan - 1);
          } else {
            if (cell.position.column - 1 < cStart) {
              columnStart = cell.position.column + cell.colSpan - 1;
              columnEnd = cell.position.column - 1;
            } else {
              columnEnd = cell.position.column - 1;
            }
          }
          needReCalc = true;
        }
      }
    }
    range = Object.assign(range, { rowStart, rowEnd, columnStart, columnEnd });

    if (needReCalc) {
      this.recalcRange(range, true);
    }
  }

  changeCellTextStyle(style: CellStyle, text: string) {}

  changeCellStyle(style: CellStyle) {
    for (let index = 0, len = this.activeArr.length; index < len; index++) {
      const range = this.activeArr[index];
      const rowStart = Math.min(range.rowStart, range.rowEnd);
      const rowEnd = Math.max(range.rowStart, range.rowEnd);
      const columnStart = Math.min(range.columnStart, range.columnEnd);
      const columnEnd = Math.max(range.columnStart, range.columnEnd);
      const cell = this.cells[rowStart][columnStart];
      // const fontWeight =
      //   this.activeCell.style.fontWeight === 'normal' ? 'bold' : 'normal';
      if (rowEnd === Infinity) {
        for (
          let i = columnStart,
            cLen = Math.min(this.columns.length - 1, columnEnd);
          i <= cLen;
          i++
        ) {
          for (const attr of Object.keys(style)) {
            this.setColumnDefaultAttr(attr, style[attr], i);
          }
          // this.setColumnDefaultAttr('fontWeight', fontWeight, i);
        }
      }
      if (columnEnd === Infinity) {
        for (
          let i = rowStart, rLen = Math.min(this.rows.length - 1, rowEnd);
          i <= rLen;
          i++
        ) {
          for (const attr of Object.keys(style)) {
            this.setRowDefaultAttr(attr, style[attr], i);
          }
          // this.setRowDefaultAttr('fontWeight', fontWeight, i);
        }
      }
      for (
        let i = rowStart, rLen = Math.min(this.rows.length - 1, rowEnd);
        i <= rLen;
        i++
      ) {
        for (
          let j = columnStart,
            cLen = Math.min(this.columns.length - 1, columnEnd);
          j <= cLen;
          j++
        ) {
          // this.cells[i][j].style.fontWeight = fontWeight;
          Object.assign(this.cells[i][j].style, style);
          if (this.cells[i][j].content.html) {
            if (style.fontSize) {
              this.cells[i][j].content.html = this.cells[i][
                j
              ].content.html.replace(
                /font-size:\s*.*;/gi,
                `font-size: ${style.fontSize}pt;`
              );
            }
            if (style.fontFamily) {
              this.cells[i][j].content.html = this.cells[i][
                j
              ].content.html.replace(
                /font-family:\s*.*;/gi,
                `font-family: ${style.fontFamily};`
              );
            }
            if (style.fontWeight) {
              this.cells[i][j].content.html = this.cells[i][
                j
              ].content.html.replace(
                /font-weight:\s*.*;/gi,
                `font-weight: ${style.fontWeight};`
              );
            }
            if (style.fontStyle) {
              this.cells[i][j].content.html = this.cells[i][
                j
              ].content.html.replace(
                /font-style:\s*.*;/gi,
                `font-style: ${style.fontStyle};`
              );
            }
            if (style.color) {
              this.cells[i][j].content.html = this.cells[i][
                j
              ].content.html.replace(
                /color:\s*.*;/gi,
                `font-style: ${style.color};`
              );
            }
          }
        }
      }
    }
    this.refreshView();
    this.canvas.focus();
  }

  onPaste = (event) => {
    // console.log('document paste');
    if (this.clipBoard) {
      this.clipBoard = null;
      return false;
    }
    const fragment = this.htmlToElement(
      event.clipboardData.getData('text/html')
    );
    // console.log(fragment);
    // console.log(event.clipboardData);
    if (fragment.querySelector('table')) {
      const trArr = fragment.querySelectorAll('table tr');
      const style = fragment.querySelector('style').innerHTML;
      const clipCells = [];
      for (let i = 0, trLen = trArr.length; i < trLen; i++) {
        const tdArr = trArr[i].querySelectorAll('td');
        if (!clipCells[i]) {
          clipCells[i] = [];
        }
        for (let j = 0, tdLen = tdArr.length; j < tdLen; j++) {
          if (clipCells[i][j]) {
            continue;
          }
          const className = tdArr[j].className;
          // this.createCell(i + this.activeCellPos);
          const cell = this.createCell(i + 1, j + 1);
          if (tdArr[j].innerText) {
            cell.content.value = tdArr[j].innerText;
          }
          if (tdArr[j].getAttribute('colSpan')) {
            cell.colSpan = parseInt(tdArr[j].getAttribute('colSpan'), 10);
          }
          if (tdArr[j].getAttribute('rowSpan')) {
            cell.rowSpan = parseInt(tdArr[j].getAttribute('rowSpan'), 10);
          }
          if (cell.colSpan > 1) {
            let col = 1;
            while (col < cell.colSpan) {
              clipCells[i][j + col] = this.createCell(i + 1, j + col + 1);
              clipCells[i][j + col].combineCell = cell;
              col++;
            }
          }
          if (cell.rowSpan > 1) {
            let row = 1;
            while (row < cell.rowSpan) {
              if (!clipCells[i + row]) {
                clipCells[i + row] = [];
              }
              clipCells[i + row][j] = this.createCell(i + row + 1, j + 1);
              clipCells[i + row][j].combineCell = cell;
              row++;
            }
          }
          if (className) {
            const background = this.getCssFromStyle(
              style,
              className,
              'background'
            );
            const fontSize = this.getCssFromStyle(
              style,
              className,
              'font-size'
            );
            const textAlign = this.getCssFromStyle(
              style,
              className,
              'text-align'
            );
            if (background) {
              cell.style.background = background;
            }
            if (fontSize) {
              cell.style.fontSize = parseFloat(fontSize);
            }
            if (textAlign) {
              cell.style.textAlign = textAlign as CanvasTextAlign;
            }
          }
          clipCells[i][j] = cell;
        }
      }
      // console.log(clipCells);
      const [clipRowStart, clipRowEnd, clipColumnStart, clipColumnEnd] = [
        0,
        clipCells.length - 1,
        0,
        clipCells[0].length - 1,
      ];
      if (!(clipRowStart === clipRowEnd && clipColumnStart === clipColumnEnd)) {
        for (let m = 0, len = this.activeArr.length; m < len; m++) {
          const range = this.activeArr[m];
          const rowStart = Math.min(range.rowStart, range.rowEnd);
          const rowEnd = rowStart + clipRowEnd - clipRowStart;
          if (rowEnd > this.rows.length - 1) {
            this.createRow(rowEnd - this.rows.length + 1);
          }
          const columnStart = Math.min(range.columnStart, range.columnEnd);
          const columnEnd = columnStart + clipColumnEnd - clipColumnStart;
          if (columnEnd > this.columns.length - 1) {
            this.createColumn(columnEnd - this.columns.length + 1);
          }
          let breakCombine = false;

          for (let i = rowStart; i <= rowEnd; i++) {
            if (
              (this.cells[i][columnStart].isCombined &&
                this.cells[i][columnStart].combineCell.position.column <
                  columnStart) ||
              (
                (this.cells[i][columnEnd].isCombined &&
                  this.cells[i][columnEnd].combineCell) ||
                this.cells[i][columnEnd]
              ).position.column +
                (
                  (this.cells[i][columnEnd].isCombined &&
                    this.cells[i][columnEnd].combineCell) ||
                  this.cells[i][columnEnd]
                ).colSpan -
                1 >
                columnEnd
            ) {
              breakCombine = true;
              break;
            }
          }
          if (!breakCombine) {
            for (let i = columnStart; i <= columnEnd; i++) {
              if (
                (this.cells[rowStart][i].isCombined &&
                  this.cells[rowStart][i].combineCell.position.row <
                    rowStart) ||
                (
                  (this.cells[rowEnd][i].isCombined &&
                    this.cells[rowEnd][i].combineCell) ||
                  this.cells[rowEnd][i]
                ).position.row +
                  (
                    (this.cells[rowEnd][i].isCombined &&
                      this.cells[rowEnd][i].combineCell) ||
                    this.cells[rowEnd][i]
                  ).rowSpan -
                  1 >
                  rowEnd
              ) {
                breakCombine = true;
                break;
              }
            }
          }
          if (breakCombine) {
            alert('cannot paste');
            return false;
          }
        }
      }
      for (let m = 0, len = this.activeArr.length; m < len; m++) {
        const range = this.activeArr[m];
        const rowStart = Math.max(1, Math.min(range.rowStart, range.rowEnd));
        const rowEnd =
          clipRowStart === clipRowEnd && clipColumnStart === clipColumnEnd
            ? Math.min(this.rows.length, Math.max(range.rowStart, range.rowEnd))
            : rowStart + clipRowEnd - clipRowStart;
        if (rowEnd > this.rows.length - 1) {
          this.createRow(rowEnd - this.rows.length + 1);
        }
        const columnStart = Math.max(
          1,
          Math.min(range.columnStart, range.columnEnd)
        );
        const columnEnd = Math.min(
          this.columns.length,
          clipRowStart === clipRowEnd && clipColumnStart === clipColumnEnd
            ? Math.min(
                this.columns.length,
                Math.max(range.columnStart, range.columnEnd)
              )
            : columnStart + clipColumnEnd - clipColumnStart
        );
        if (columnEnd > this.columns.length - 1) {
          this.createColumn(columnEnd - this.columns.length + 1);
        }
        for (let i = rowStart; i <= rowEnd; i++) {
          for (let j = columnStart; j <= columnEnd; j++) {
            const clipCell =
              clipRowStart === clipRowEnd && clipColumnStart === clipColumnEnd
                ? clipCells[0][0]
                : clipCells[i - rowStart][j - columnStart];
            if (!clipCell) {
              continue;
            }
            if (
              this.cells[i][j].isCombined ||
              this.cells[i][j].rowSpan > 1 ||
              this.cells[i][j].colSpan > 1
            ) {
              this.unCombine(
                (this.cells[i][j].isCombined && this.cells[i][j].combineCell) ||
                  this.cells[i][j]
              );
            }
            this.cells[i][j].style = clipCell.style;
            this.cells[i][j].content = clipCell.content;
            // this.cells[i][j].isCombined = clipCell.isCombined;
            this.cells[i][j].rowSpan = clipCell.rowSpan;
            this.cells[i][j].colSpan = clipCell.colSpan;
            this.cells[i][j].combineCell =
              (clipCell.combineCell &&
                this.cells[rowStart + clipCell.combineCell.position.row - 1][
                  columnStart + clipCell.combineCell.position.column - 1
                ]) ||
              null;
          }
        }

        // if (this.clipBoard.isCut) {
        //   for (let i = clipRowStart; i <= clipRowEnd; i++) {
        //     for (let j = clipColumnStart; j <= clipColumnEnd; j++) {
        //       if (
        //         inRange(i, rowStart, rowEnd, true) &&
        //         inRange(j, columnStart, columnEnd, true)
        //       ) {
        //         continue;
        //       }
        //       this.cells[i][j] = this.createCell(i, j);
        //     }
        //   }
        //   // this.clipBoard = null;
        // } else if (
        //   (inRange(rowStart, clipRowStart, clipRowEnd, true) ||
        //     inRange(rowEnd, clipRowStart, clipRowEnd, true)) &&
        //   (inRange(columnStart, clipColumnStart, clipColumnEnd, true) ||
        //     inRange(columnEnd, clipColumnStart, clipColumnEnd))
        // ) {
        //   // this.clipBoard = null;
        // }

        this.activeArr[m] = { rowStart, rowEnd, columnStart, columnEnd };
      }

      this.resetCellPerspective(this.activeCell);
    } else if (event.clipboardData.getData('text/plain')) {
      const value = event.clipboardData.getData('text/plain');
      for (let i = 0, len = this.activeArr.length; i < len; i++) {
        const rowStart = Math.max(
          Math.min(this.activeArr[i].rowEnd, this.activeArr[i].rowStart),
          this.cells[1][0].position.row
        );
        const rowEnd = Math.min(
          Math.max(this.activeArr[i].rowEnd, this.activeArr[i].rowStart),
          this.cells[this.rows.length - 1][0].position.row
        );
        const columnStart = Math.max(
          Math.min(this.activeArr[i].columnStart, this.activeArr[i].columnEnd),
          this.cells[0][1].position.column
        );
        const columnEnd = Math.min(
          Math.max(this.activeArr[i].columnStart, this.activeArr[i].columnEnd),
          this.cells[0][this.columns.length - 1].position.column
        );
        for (let m = rowStart; m <= rowEnd; m++) {
          for (let n = columnStart; n <= columnEnd; n++) {
            this.cells[m][n].content = { value };
          }
        }
      }
      this.resetCellPerspective(this.activeCell);
    } else if (
      event.clipboardData.items &&
      Array.from(event.clipboardData.items).find(
        (item: any) => item.kind === 'file' && /image\/\w*/.test(item.type)
      )
    ) {
      for (const item of event.clipboardData.items) {
        if (item.kind === 'file' && /image\/\w*/.test(item.type)) {
          createImageBitmap(item.getAsFile()).then((img) => {
            this.addImage(img);
          });
          return;
        }
      }
    }
    this.canvas.focus();
  };

  addImage(img) {
    const x =
      this.floatArr.length && !this.activeCell
        ? this.floatArr[this.floatArr.length - 1].x + this.offsetLeft
        : this.activeCell.x + this.activeCell.width / 2;
    const y =
      this.floatArr.length && !this.activeCell
        ? this.floatArr[this.floatArr.length - 1].y + this.offsetTop
        : this.activeCell.y + this.activeCell.height / 2;
    const floatElement = new FloatElement(x, y, img.width, img.height, img);
    this.floatArr.forEach((elem) => (elem.isActive = false));
    this.floatArr.push(floatElement);
    // this.floatCtx.drawImage(img, floatElement.x, floatElement.y);
    this.drawFloat();
    this.activeArr = [];
    floatElement.isActive = true;
    this.resetCellPerspective(this.activeCell);
    this.activeCellPos = null;
  }

  toggleItalic() {}

  getCellDefaultAttr(attr: string, row: number, column: number) {
    // return this.rows[row][name].lastModifyTime >
    //   this.columns[column][name].lastModifyTime
    //   ? this.rows[row][name].value
    //   : this.columns[column][name].value;
    return (
      (this.rows[row] && this.rows[row].style && this.rows[row].style[attr]) ||
      (this.columns[column] &&
        this.columns[column].style &&
        this.columns[column].style[attr])
    );
  }
  setRowDefaultAttr(attr: string, value: any, row: number) {
    // this.rows[row][name].value = value;
    // this.rows[row][name].lastModifyTime = new Date().getTime();
    if (!this.rows[row].style) {
      this.rows[row].style = {};
    }
    this.rows[row].style[attr] = value;
  }

  setColumnDefaultAttr(attr: string, value: any, column: number) {
    // this.columns[column][name].value = value;
    // this.columns[column][name].lastModifyTime = new Date().getTime();
    if (!this.columns[column].style) {
      this.columns[column].style = {};
    }
    this.columns[column].style[attr] = value;
  }

  htmlToElement(html: string) {
    const template = document.createElement('template');
    html = html.trim();
    template.innerHTML = html;
    return template.content;
  }

  getCssFromStyle(style: string, className: string, attr: string) {
    const pattern = `\\.${className}\\s*{\\s*[^}]*${attr}:([^;}]+);`;
    const result = new RegExp(pattern).exec(style);
    if (result) {
      return result[1];
    } else {
      return null;
    }
  }

  getViewFloatElementByPoint(pointX: number, pointY: number) {
    for (let i = this.floatArr.length - 1; i >= 0; i--) {
      const x = this.floatArr[i].x;
      const y = this.floatArr[i].y;
      const width = this.floatArr[i].width;
      const height = this.floatArr[i].height;
      const isActive = this.floatArr[i].isActive;
      const radius = this.style.activeFloatElementResizeArcRadius;
      if (
        !inRange(
          x - this.scrollLeft - radius,
          this.offsetLeft,
          this.offsetLeft + this.clientWidth
        ) &&
        !inRange(
          x + width - this.scrollLeft + radius,
          this.offsetLeft,
          this.offsetLeft + this.clientWidth
        ) &&
        !inRange(
          y - this.scrollTop - radius,
          this.offsetTop,
          this.offsetTop + this.clientHeight
        ) &&
        !inRange(
          y + height - this.scrollTop + radius,
          this.offsetTop,
          this.offsetTop + this.clientHeight
        )
      ) {
        continue;
      }
      if (
        inRange(pointX + this.scrollLeft, x - radius, x + width + radius) &&
        inRange(pointY + this.scrollTop, y - radius, y + height + radius)
      ) {
        return this.floatArr[i];
      }
    }
    return null;
  }

  getViewCellByPoint(pointX: number, pointY: number) {
    for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
      for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
        const cell = this.viewCells[i][j].isCombined
          ? this.viewCells[i][j].combineCell
          : this.viewCells[i][j];
        if (this.inCellArea(pointX, pointY, cell)) {
          return cell;
        }
      }
    }
    return null;
  }

  clearRect(ctx: CanvasRenderingContext2D) {
    ctx.clearRect(
      0,
      0,
      this.width / this.multiple,
      this.height / this.multiple
    );
  }

  setTransform() {
    const scale = this.multiple;
    this.ctx.transform(scale, 0, 0, scale, 0, 0);
    this.actionCtx.transform(scale, 0, 0, scale, 0, 0);
    this.animationCtx.transform(scale, 0, 0, scale, 0, 0);
    this.floatCtx.transform(scale, 0, 0, scale, 0, 0);
    this.offscreenCtx.transform(scale, 0, 0, scale, 0, 0);
    this.floatActionCtx.transform(scale, 0, 0, scale, 0, 0);
  }

  filterOffscreenFloat(floatArr) {
    return floatArr.filter((elem) => {
      const x = elem.x;
      const y = elem.y;
      const width = elem.width;
      const height = elem.height;

      const radius = this.style.activeFloatElementResizeArcRadius;
      return !(
        !inRange(
          x - this.scrollLeft - radius,
          this.offsetLeft,
          this.offsetLeft + this.clientWidth
        ) &&
        !inRange(
          x + width - this.scrollLeft + radius,
          this.offsetLeft,
          this.offsetLeft + this.clientWidth
        ) &&
        !inRange(
          y - this.scrollTop - radius,
          this.offsetTop,
          this.offsetTop + this.clientHeight
        ) &&
        !inRange(
          y + height - this.scrollTop + radius,
          this.offsetTop,
          this.offsetTop + this.clientHeight
        )
      );
    });
  }

  getFloatElementResizePos(elem: FloatElement, pointX: number, pointY: number) {
    const radius = this.style.activeFloatElementResizeArcRadius;
    const x = elem.x;
    const y = elem.y;
    const width = elem.width;
    const height = elem.height;
    let pos = LogicPosition.Other;
    if (
      inRange(
        pointX + this.scrollLeft,
        x + width - radius,
        x + width + radius
      ) &&
      inRange(pointY + this.scrollTop, y + height - radius, y + height + radius)
    ) {
      pos = LogicPosition.RightBottom;
    } else if (
      inRange(pointX + this.scrollLeft, x - radius, x + radius) &&
      inRange(pointY + this.scrollTop, y - radius, y + radius)
    ) {
      pos = LogicPosition.LeftTop;
    } else if (
      inRange(
        pointX + this.scrollLeft,
        x + width / 2 - radius,
        x + width / 2 + radius
      ) &&
      inRange(pointY + this.scrollTop, y - radius, y + radius)
    ) {
      pos = LogicPosition.Top;
    } else if (
      inRange(
        pointX + this.scrollLeft,
        x + width - radius,
        x + width + radius
      ) &&
      inRange(pointY + this.scrollTop, y - radius, y + radius)
    ) {
      pos = LogicPosition.rightTop;
    } else if (
      inRange(pointX + this.scrollLeft, x - radius, x + radius) &&
      inRange(
        pointY + this.scrollTop,
        y + height / 2 - radius,
        y + height / 2 + radius
      )
    ) {
      pos = LogicPosition.Left;
    } else if (
      inRange(
        pointX + this.scrollLeft,
        x + width - radius,
        x + width + radius
      ) &&
      inRange(
        pointY + this.scrollTop,
        y + height / 2 - radius,
        y + height / 2 + radius
      )
    ) {
      pos = LogicPosition.Right;
    } else if (
      inRange(pointX + this.scrollLeft, x - radius, x + radius) &&
      inRange(pointY + this.scrollTop, y + height - radius, y + height + radius)
    ) {
      pos = LogicPosition.LeftBottom;
    } else if (
      inRange(
        pointX + this.scrollLeft,
        x + width / 2 - radius,
        x + width / 2 + radius
      ) &&
      inRange(pointY + this.scrollTop, y + height - radius, y + height + radius)
    ) {
      pos = LogicPosition.Bottom;
    }
    return pos;
  }
}
