import { Style } from 'src/app/core/model/style.const';
import { CellRange } from 'src/app/core/model/cell-range.class';
import { inRange } from 'src/app/core/decorator/utils/function';
import { KeyCode } from 'src/app/core/model/key-code.enmu';
import { Cell } from './cell.class';

export class Panel {
  width = 0;
  height = 0;
  viewRowCount = 0;
  viewColumnCount = 0;
  offsetLeft = 60;
  offsetTop = 30;
  columns = [];
  rows = [];
  cells: Cell[][] = [];
  viewCells: Cell[][] = [];
  editingCell: Cell;
  scrollBarWidth = Style.scrollBarWidth;
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
  unActiveCellPos: any;
  //   activeCellPos: any = { row: 1, column: 1, rangeIndex: 0 };
  private _activeCellPos: any;
  activeCell: any;
  get activeCellPos() {
    return this._activeCellPos;
  }
  set activeCellPos(val) {
    this._activeCellPos = val;
    this.activeCell = this.cells[val.row][val.column].isCombined
      ? this.cells[val.row][val.column].combineCell
      : this.cells[val.row][val.column];
  }

  activeArr: CellRange[] = [
    { rowStart: 1, columnStart: 1, rowEnd: 1, columnEnd: 1 },
  ];
  unActiveRange: CellRange;
  resizeColumnCell: Cell;
  resizeRowCell: Cell;

  canvas: HTMLCanvasElement;
  actionCanvas: HTMLCanvasElement;

  ctx: CanvasRenderingContext2D;
  actionCtx: CanvasRenderingContext2D;

  init() {
    this.ctx = this.canvas.getContext('2d');
    this.actionCtx = this.actionCanvas.getContext('2d');
    this.resize();
    this.activeCellPos = { row: 1, column: 1, rangeIndex: 0 };
    this.refreshView();
    this.canvas.focus();
  }

  resize() {
    this.width = this.canvas.offsetWidth;
    this.height = this.canvas.offsetHeight;
    this.clientWidth = this.width - this.offsetLeft - Style.scrollBarWidth;
    this.clientHeight = this.height - this.offsetTop - Style.scrollBarWidth;
    this.canvas.width = this.width;
    this.canvas.height = this.height;
    this.actionCanvas.width = this.width;
    this.actionCanvas.height = this.height;
    this.viewRowCount =
      Math.ceil((this.height - this.offsetTop) / Style.cellHeight) + 2;
    if (this.viewCells.length < this.viewRowCount) {
      this.createRow(this.viewRowCount - this.viewCells.length);
    }
    this.viewColumnCount =
      Math.ceil((this.width - this.offsetLeft) / Style.cellWidth) + 2;
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

    this.canvas.width = this.width;
    this.canvas.height = this.height;
    this.actionCanvas.width = this.width;
    this.actionCanvas.height = this.height;
  }

  refreshView() {
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

    const canvas = document.createElement('canvas');
    canvas.width = this.width;
    canvas.height = this.height;
    const ctx = canvas.getContext('2d');
    this.setActive();
    this.drawPanel(ctx);
    this.drawRuler(ctx);
    this.drawScrollBar(ctx);
    this.ctx.clearRect(0, 0, this.width, this.height);
    this.ctx.drawImage(canvas, 0, 0);
  }

  createCell(rk: number, ck: number): Cell {
    const isXRuler = rk === 0;
    const isYRuler = ck === 0;
    return {
      columns: this.columns,
      rows: this.rows,
      position: { row: rk, column: ck },
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
      get width() {
        let colSpan = this.colSpan;
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
        let rowSpan = this.rowSpan;
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
      rowSpan: 1,
      colSpan: 1,
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
      this.width - Style.scrollBarWidth - this.offsetLeft,
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
    ctx.fillRect(
      this.width - Style.scrollBarWidth,
      this.height - Style.scrollBarWidth,
      this.width,
      this.height
    );
    ctx.strokeRect(
      this.width - Style.scrollBarWidth,
      this.offsetTop,
      this.width,
      this.height - Style.scrollBarWidth - this.offsetTop
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
    ctx.clearRect(0, 0, this.width, this.offsetTop);
    ctx.clearRect(0, 0, this.offsetLeft, this.height);
    const columns = this.viewCells[0].slice(1);
    const rows = this.viewCells.slice(1).map((cells) => cells[0]);
    ctx.textAlign = columns[0].textAlign as CanvasTextAlign;
    ctx.textBaseline = columns[0].textBaseline as CanvasTextBaseline;
    ctx.strokeStyle = Style.rulerCellBorderColor;
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
      if (columns[i].width === 0) {
        continue;
      }
      const x =
        columns[i].x -
        this.scrollLeft +
        (columns[i - 1] && !columns[i - 1].width
          ? Style.rulerResizeGapWidth / 2
          : 0);
      const width =
        columns[i].width -
        (columns[i + 1] && !columns[i + 1].width
          ? Style.rulerResizeGapWidth / 2
          : 0) -
        (columns[i - 1] && !columns[i - 1].width
          ? Style.rulerResizeGapWidth / 2
          : 0);
      ctx.fillRect(x, columns[i].y, width, columns[i].height);
      ctx.strokeRect(x, columns[i].y, width, columns[i].height);
      if (columns[i].content.value) {
        ctx.fillStyle = columns[0].color;
        ctx.fillText(
          columns[i].content.value,
          x + width / 2,
          columns[i].y + columns[i].height / 2,
          width - 2 * columns[0].borderWidth
        );
      }
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
      if (rows[i].height === 0) {
        continue;
      }
      const y =
        rows[i].y -
        this.scrollTop +
        (rows[i - 1] && !rows[i - 1].height
          ? Style.rulerResizeGapWidth / 2
          : 0);
      const height =
        rows[i].height -
        (rows[i + 1] && !rows[i + 1].height
          ? Style.rulerResizeGapWidth / 2
          : 0) -
        (rows[i - 1] && !rows[i - 1].height
          ? Style.rulerResizeGapWidth / 2
          : 0);
      ctx.fillRect(rows[i].x, y, rows[i].width, height);
      ctx.strokeRect(rows[i].x, y, rows[i].width, height);
      if (rows[i].content.value) {
        ctx.fillStyle = columns[0].color;
        ctx.fillText(
          rows[i].content.value,
          rows[i].x + rows[i].width / 2,
          y + height / 2,
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

  drawCell(ctx: CanvasRenderingContext2D, drawCell: Cell) {
    const cell = drawCell.isCombined ? drawCell.combineCell : drawCell;
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
            cell.x - this.scrollLeft + 3 * Style.cellBorderWidth,
            cell.y - this.scrollTop + 3 * Style.cellBorderWidth,
            cell.width - 6 * Style.cellBorderWidth,
            cell.height - 6 * Style.cellBorderWidth
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
              (this.columns[i].width ? Style.rulerResizeGapWidth : 0),
            this.columns[i].x -
              this.scrollLeft +
              this.columns[i].width +
              (this.columns[i + 1] && this.columns[i + 1].width
                ? Style.rulerResizeGapWidth
                : 0)
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
    // this.panel.nativeElement.focus();
    event.preventDefault();
    event.stopPropagation();
    this.canvas.focus();
    event.preventDefault();
    // console.log('mousedown', event);

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
      x: event.offsetX,
      y: event.offsetY,
      lastModifyTime: new Date().getTime(),
    };
    if (this.inSelectAllArea(event.offsetX, event.offsetY)) {
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
    } else if (this.inRulerXArea(event.offsetX, event.offsetY)) {
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
        // if (ctx.isPointInPath(event.offsetX, event.offsetY)) {
        if (
          this.viewCells[0][i].width >
            (i === 1 ? 1 : 2) * Style.rulerResizeGapWidth &&
          inRange(
            event.x,
            this.viewCells[0][i].x -
              this.scrollLeft +
              (i === 1 || !this.viewCells[0][i].width
                ? 0
                : Style.rulerResizeGapWidth),
            this.viewCells[0][i].x -
              this.scrollLeft +
              this.viewCells[0][i].width -
              (i === 1 ? 1 : this.viewCells[0][i].width ? 2 : 0) *
                Style.rulerResizeGapWidth
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
          this.viewCells[0][i].width <=
            (i === 1 ? 1 : 2) * Style.rulerResizeGapWidth ||
          inRange(
            event.offsetX,
            this.viewCells[0][i].x -
              this.scrollLeft +
              this.viewCells[0][i].width -
              (this.viewCells[0][i].width ? Style.rulerResizeGapWidth : 0),
            this.viewCells[0][i].x -
              this.scrollLeft +
              this.viewCells[0][i].width +
              (this.viewCells[0][i + 1] && this.viewCells[0][i + 1].width
                ? Style.rulerResizeGapWidth
                : 0)
          )
        ) {
          this.resizeColumnCell = this.viewCells[0][i];
          this.state.isResizeColumn = true;
          break;
        }
      }
    } else if (this.inRulerYArea(event.offsetX, event.offsetY)) {
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
        // if (ctx.isPointInPath(event.offsetX, event.offsetY)) {
        if (
          rowCells[i].height > (i === 1 ? 1 : 2) * Style.rulerResizeGapWidth &&
          inRange(
            event.offsetY,
            rowCells[i].y -
              this.scrollTop +
              (i === 1 || !rowCells[i].width ? 0 : Style.rulerResizeGapWidth),
            rowCells[i].y -
              this.scrollTop +
              rowCells[i].height -
              (i === 1 ? 1 : rowCells[i].width ? 2 : 0) *
                Style.rulerResizeGapWidth
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
          rowCells[i].height <= (i === 1 ? 1 : 2) * Style.rulerResizeGapWidth ||
          inRange(
            event.offsetY,
            rowCells[i].y -
              this.scrollTop +
              rowCells[i].height -
              (rowCells[i].height ? Style.rulerResizeGapWidth : 0),
            rowCells[i].y -
              this.scrollTop +
              rowCells[i].height +
              (rowCells[i + 1] && rowCells[i + 1].height
                ? Style.rulerResizeGapWidth
                : 0)
          )
        ) {
          console.log('dblclick', isDblClick, event);
          if (isDblClick) {
            this.resizeRow(
              rowCells[i].position.row,
              Style.cellHeight - rowCells[i].height
            );
            break;
          } else {
            this.resizeRowCell = rowCells[i];
            this.state.isResizeRow = true;
            break;
          }
        }
      }
    } else if (this.inScrollXBarArea(event.offsetX, event.offsetY)) {
      console.log('scrollx');
      if (this.inThumbAreaOfScrollBarX(event.offsetX, event.offsetY, true)) {
        this.state.isSelectScrollXThumb = true;
      } else if (
        inRange(
          event.offsetX,
          this.offsetLeft,
          this.getScrollXThumbLeft(this.getScrollXThumbHeight())
        )
      ) {
        this.scrollX(-1 * this.viewColumnCount * Style.cellWidth);
        this.setActive();
        this.drawRuler(this.ctx);
      } else {
        this.scrollX(this.viewColumnCount * Style.cellWidth);
        this.setActive();
        this.drawRuler(this.ctx);
      }
    } else if (this.inScrollYBarArea(event.offsetX, event.offsetY)) {
      console.log('scrolly');
      if (this.inThumbAreaOfScrollBarY(event.offsetX, event.offsetY, true)) {
        this.state.isSelectScrollYThumb = true;
      } else if (
        inRange(
          event.offsetY,
          this.offsetTop,
          this.getScrollYThumbTop(this.getScrollYThumbHeight())
        )
      ) {
        this.scrollY(-1 * this.viewRowCount * Style.cellHeight);
        this.setActive();
        this.drawRuler(this.ctx);
      } else {
        this.scrollY(this.viewRowCount * Style.cellHeight);
        this.setActive();
        this.drawRuler(this.ctx);
      }
    } else if (this.inCellsArea(event.offsetX, event.offsetY)) {
      // const canvas = document.createElement('canvas');
      // canvas.width = this.width;
      // canvas.height = this.height;
      // const ctx = canvas.getContext('2d');
      for (let rLen = this.viewCells.length, i = rLen - 1; i > 0; i--) {
        for (let cLen = this.viewCells[i].length, j = cLen - 1; j > 0; j--) {
          const cell = this.viewCells[i][j].isCombined
            ? this.viewCells[i][j].combineCell
            : this.viewCells[i][j];
          // ctx.clearRect(0, 0, this.width, this.height);
          // ctx.rect(
          //   cell.x - this.scrollLeft,
          //   cell.y - this.scrollTop,
          //   cell.width,
          //   cell.height
          // );
          // if (ctx.isPointInPath(event.offsetX, event.offsetY)) {
          if (this.inCellArea(event.offsetX, event.offsetY, cell)) {
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
                this.editingCell = cell;
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
                  rowEnd: cell.position.row + cell.rowSpan - 1,
                  columnEnd: cell.position.column + cell.colSpan - 1,
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
      (!this.inCellsArea(event.offsetX, event.offsetY) &&
        !this.state.isSelectCell) ||
      this.state.isSelectScrollYThumb ||
      this.state.isSelectScrollXThumb ||
      this.state.isResizeColumn
    ) {
      if (
        this.state.isResizeColumn ||
        (!this.state.isSelectRulerX &&
          !this.state.isSelectRulerY &&
          this.inRulerXResizeGap(event.offsetX, event.offsetY))
      ) {
        this.canvas.style.cursor = 'col-resize';
      } else if (
        this.state.isResizeRow ||
        (!this.state.isSelectRulerX &&
          !this.state.isSelectRulerY &&
          this.inRulerYResizeGap(event.offsetX, event.offsetY))
      ) {
        this.canvas.style.cursor = 'row-resize';
      } else {
        this.canvas.style.cursor = 'default';
      }
    } else {
      this.canvas.style.cursor = 'cell';
    }
    const preIsScrollXThumbHover = this.state.isScrollXThumbHover;
    const preIsScrollYThumbHover = this.state.isScrollYThumbHover;
    this.state.isScrollXThumbHover = this.inThumbAreaOfScrollBarX(
      event.offsetX,
      event.offsetY
    );
    this.state.isScrollYThumbHover = this.inThumbAreaOfScrollBarY(
      event.offsetX,
      event.offsetY
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
          range = this.calcActive(event.offsetX, event.offsetY);
        } else if (this.state.unSelectCell) {
          range = this.calcUnActive(event.offsetX, event.offsetY);
        } else if (this.state.isSelectScrollYThumb) {
          this.calcScrollY(event.offsetX, event.offsetY);
        } else if (this.state.isSelectScrollXThumb) {
          this.calcScrollX(event.offsetX, event.offsetY);
        } else if (this.state.isSelectRulerX) {
          range = this.calcAcitiveRulerX(event.offsetX, event.offsetY);
        } else if (this.state.isSelectRulerY) {
          range = this.calcActiveRulerY(event.offsetX, event.offsetY);
        } else if (this.state.unSelectRulerX) {
          range = this.calcUnActiveRulerX(event.offsetX, event.offsetY);
        } else if (this.state.unSelectRulerY) {
          range = this.calcUnActiveRulerY(event.offsetX, event.offsetY);
        } else if (this.state.isResizeColumn) {
          this.resizeColumn(
            this.resizeColumnCell.position.column,
            event.offsetX - this.mousePoint.x
          );
        } else if (this.state.isResizeRow) {
          this.resizeRow(
            this.resizeRowCell.position.row,
            event.offsetY - this.mousePoint.y
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

        this.mousePoint = { x: event.offsetX, y: event.offsetY };
        this.isTicking = false;
      });
    }

    this.isTicking = true;
  }

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

  resizeRow(row: number, deltaHeight: number) {
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
          const rowReverse = cell.position.row < this.activeCellPos.row;
          const columnReverse =
            cell.position.column < this.activeCellPos.column;
          const [rStart, rEnd, cStart, cEnd] = [
            Math.min(this.activeCellPos.row, cell.position.row),
            Math.max(this.activeCellPos.row, cell.position.row),
            Math.min(this.activeCellPos.column, cell.position.column),
            Math.max(this.activeCellPos.column, cell.position.column),
          ];
          let [rowStart, rowEnd, columnStart, columnEnd] = [
            rStart,
            rEnd,
            cStart,
            cEnd,
          ];
          for (let m = rStart; m <= rEnd; m++) {
            for (let n = cStart; n <= cEnd; n++) {
              const activeCell = this.cells[m][n].isCombined
                ? this.cells[m][n].combineCell
                : this.cells[m][n];
              if (activeCell.position.row < rStart) {
                rowStart = activeCell.position.row;
              }
              if (activeCell.position.row + activeCell.rowSpan - 1 > rEnd) {
                rowEnd = activeCell.position.row + activeCell.rowSpan - 1;
              }
              if (activeCell.position.column < cStart) {
                columnStart = activeCell.position.column;
              }
              if (activeCell.position.column + activeCell.colSpan - 1 > cEnd) {
                columnEnd = activeCell.position.column + activeCell.colSpan - 1;
              }
            }
          }
          console.log(rowReverse, columnReverse ,rowStart, rowEnd, columnStart, columnEnd)
          return {
            rowStart: rowReverse ? rowEnd : rowStart,
            rowEnd: rowReverse ? rowStart : rowEnd,
            columnStart: columnReverse ? columnEnd : columnStart,
            columnEnd: columnReverse ? columnStart : columnEnd,
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
          //   return {
          //     rowStart: this.unActiveRange.rowStart,
          //     columnStart: this.unActiveRange.columnStart,
          //     rowEnd: cell.position.row,
          //     columnEnd: cell.position.column,
          //   };
          const rowReverse = cell.position.row < this.activeCellPos.row;
          const columnReverse =
            cell.position.column < this.activeCellPos.column;
          if (this.inCellArea(x, y, cell)) {
            const [rStart, rEnd, cStart, cEnd] = [
              Math.min(this.unActiveCellPos.row, cell.position.row),
              Math.max(this.unActiveCellPos.row, cell.position.row),
              Math.min(this.unActiveCellPos.column, cell.position.column),
              Math.max(this.unActiveCellPos.column, cell.position.column),
            ];
            let [rowStart, rowEnd, columnStart, columnEnd] = [
              rStart,
              rEnd,
              cStart,
              cEnd,
            ];
            for (let m = rStart; m <= rEnd; m++) {
              for (let n = cStart; n <= cEnd; n++) {
                const activeCell = this.cells[m][n].isCombined
                  ? this.cells[m][n].combineCell
                  : this.cells[m][n];
                if (activeCell.position.row < rStart) {
                  rowStart = activeCell.position.row;
                }
                if (activeCell.position.row + activeCell.rowSpan - 1 > rEnd) {
                  rowEnd = activeCell.position.row + activeCell.rowSpan - 1;
                }
                if (activeCell.position.column < cStart) {
                  columnStart = activeCell.position.column;
                }
                if (
                  activeCell.position.column + activeCell.colSpan - 1 >
                  cEnd
                ) {
                  columnEnd =
                    activeCell.position.column + activeCell.colSpan - 1;
                }
              }
            }
            return {
              rowStart: rowReverse ? rowEnd : rowStart,
              rowEnd: rowReverse ? rowStart : rowEnd,
              columnStart: columnReverse ? columnEnd : columnStart,
              columnEnd: columnReverse ? columnStart : columnEnd,
            };
          }
        }
      }
    }
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
    console.log(rowStart, rowEnd, columnStart, columnEnd);

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
    console.log(this.activeArr);
  }

  createColumn(count: number) {
    for (let i = 0; i < count; i++) {
      const columnLength = this.columns.length;
      this.columns.push({
        x: columnLength
          ? this.columns[columnLength - 1].x +
            this.columns[columnLength - 1].width
          : 0,
        width: columnLength ? Style.cellWidth : this.offsetLeft,
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
            Style.cellWidth
        ) + 1
      );
      // for (
      //   let i = 0,
      //     len =
      //       Math.ceil(
      //         (this.scrollLeft - this.scrollWidth + this.clientWidth) /
      //           Style.cellWidth
      //       ) + 1;
      //   i < len;
      //   i++
      // ) {
      //   const columnLength = this.columns.length;
      //   this.columns.push({
      //     x:
      //       this.columns[columnLength - 1].x +
      //       this.columns[columnLength - 1].width,
      //     width: Style.cellWidth,
      //   });
      //   this.cells.forEach((row, rk) =>
      //     row.push(this.createCell(rk, columnLength))
      //   );
      // }
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

  createRow(count: number = 1) {
    for (let i = 0; i < count; i++) {
      this.rows.push({
        y: this.rows.length
          ? this.rows[this.rows.length - 1].y +
            this.rows[this.rows.length - 1].height
          : 0,
        height: this.rows.length ? Style.cellHeight : this.offsetTop,
      });
      this.cells.push(
        Array.from({ length: this.columns.length }).map((cv, ck) =>
          this.createCell(this.cells.length, ck)
        )
      );
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
            Style.cellHeight
        ) + 1
      );
      // for (
      //   let i = 0,
      //     len =
      //       Math.ceil(
      //         (this.scrollTop - this.scrollHeight + this.clientHeight) /
      //           Style.cellHeight
      //       ) + 1;
      //   i < len;
      //   i++
      // ) {
      //   this.rows.push({
      //     y:
      //       this.rows[this.rows.length - 1].y +
      //       this.rows[this.rows.length - 1].height,
      //     height: Style.cellHeight,
      //   });
      //   this.cells.push(
      //     Array.from({ length: this.cells[0].length }).map((cv, ck) =>
      //       this.createCell(this.cells.length, ck)
      //     )
      //   );
      // }
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
      } else if (event.code === KeyCode.ArrowUp) {
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
          const isExpand = range.rowEnd >= range.rowStart;
          range.rowEnd =
            range.rowEnd + 1 > this.rows.length - 1 || event.ctrlKey
              ? this.rows.length - 1
              : range.rowEnd + 1;
          this.recalcRange(range, isExpand);
        } else if (event.code === KeyCode.ArrowUp) {
          const isExpand = range.rowEnd <= range.rowStart;
          range.rowEnd =
            range.rowEnd - 1 < 1 || event.ctrlKey ? 1 : range.rowEnd - 1;
          this.recalcRange(range, isExpand);
        }
        // const [rStart, rEnd, cStart, cEnd] = [
        //   Math.min(range.rowStart, range.rowEnd),
        //   Math.max(range.rowStart, range.rowEnd),
        //   Math.min(range.columnStart, range.columnEnd),
        //   Math.max(range.columnStart, range.columnEnd),
        // ];
        // let [rowStart, rowEnd, columnStart, columnEnd] = [
        //   rStart,
        //   rEnd,
        //   cStart,
        //   cEnd,
        // ];
        // for (let i = rStart; i <= rEnd; i++) {
        //   for (
        //     let j = cStart, cLen = Math.min(cEnd, this.columns.length - 1);
        //     j <= cLen;
        //     j++
        //   ) {
        //     if (this.cells[i][j].hidden) {
        //       continue;
        //     }
        //     const cell = this.cells[i][j].isCombined
        //       ? this.cells[i][j].combineCell
        //       : this.cells[i][j];
        //     if (cell.position.row < rStart) {
        //       rowStart = cell.position.row;
        //     }
        //     if (cell.position.row + cell.rowSpan - 1 > rEnd) {
        //       rowEnd = cell.position.row + cell.rowSpan - 1;
        //     }
        //     if (cell.position.column < cStart) {
        //       columnStart = cell.position.column;
        //     }
        //     if (cell.position.column + cell.colSpan - 1 > cEnd) {
        //       columnEnd = cell.position.column + cell.colSpan - 1;
        //     }
        //   }
        // }
        // this.activeArr[this.activeCellPos.rangeIndex] = {
        //   rowStart,
        //   rowEnd,
        //   columnStart,
        //   columnEnd,
        // };
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
      } else if (event.code === KeyCode.ArrowLeft) {
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
          const isExpand = range.columnEnd >= range.columnStart;
          range.columnEnd =
            range.columnEnd + 1 > this.cells[0].length - 1 || event.ctrlKey
              ? this.cells[0].length - 1
              : range.columnEnd + 1;
          this.recalcRange(range, isExpand);
        } else if (event.code === KeyCode.ArrowLeft) {
          const isExpand = range.columnEnd <= range.columnStart;
          range.columnEnd =
            range.columnEnd - 1 < 1 || event.ctrlKey ? 1 : range.columnEnd - 1;
          this.recalcRange(range, isExpand);
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
        if (event.code === KeyCode.Tab) {
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
        } else if (event.code === KeyCode.Enter) {
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
        if (event.code === KeyCode.Tab) {
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
        } else if (event.code === KeyCode.Enter) {
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
      if (!this.isTicking) {
        requestAnimationFrame(() => {
          this.resetCellPerspective(
            this.cells[this.activeCellPos.row][this.activeCellPos.column]
          );
          this.isTicking = false;
        });
      }
    } else {
      //   let range = this.activeArr[this.activeCellPos.rangeIndex];
      let { row, column, rangeIndex } = this.activeCellPos;
      let next = false;
      if (event.code === KeyCode.Tab) {
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

        // if (
        //   event.shiftKey
        //     ? this.activeCell.position.column - 1 <
        //       Math.min(range.columnStart, range.columnEnd)
        //     : this.activeCell.position.column + this.activeCell.colSpan >
        //       Math.max(range.columnStart, range.columnEnd)
        // ) {
        //   if (
        //     event.shiftKey
        //       ? this.activeCellPos.row - 1 <
        //         Math.min(range.rowStart, range.rowEnd)
        //       : this.activeCellPos.row + 1 >
        //         Math.max(range.rowStart, range.rowEnd)
        //   ) {
        //     const rangeIndex = event.shiftKey
        //       ? this.activeCellPos.rangeIndex - 1 < 0
        //         ? this.activeArr.length - 1
        //         : this.activeCellPos.rangeIndex - 1
        //       : this.activeCellPos.rangeIndex + 1 > this.activeArr.length - 1
        //       ? 0
        //       : this.activeCellPos.rangeIndex + 1;
        //     range = this.activeArr[rangeIndex];
        //     this.activeCellPos = {
        //       row: event.shiftKey
        //         ? Math.max(range.rowStart, range.rowEnd)
        //         : Math.min(range.rowStart, range.rowEnd),
        //       column: event.shiftKey
        //         ? Math.max(range.columnStart, range.columnEnd)
        //         : Math.min(range.columnStart, range.columnEnd),
        //       rangeIndex,
        //     };
        //   } else {
        //     event.shiftKey
        //       ? this.activeCellPos.row--
        //       : this.activeCellPos.row++;
        //     this.activeCellPos.column = event.shiftKey
        //       ? Math.max(range.columnStart, range.columnEnd)
        //       : (this.activeCellPos.column = Math.min(
        //           range.columnStart,
        //           range.columnEnd
        //         ));
        //   }
        // } else {
        //   event.shiftKey
        //     ? this.activeCellPos.column--
        //     : this.activeCellPos.column++;
        // }
      } else if (event.code === KeyCode.Enter) {
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
              //   const cell =
              //     (this.cells[row][column].isCombined &&
              //       this.cells[row][column].combineCell) ||
              //     this.cells[row][column];
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
        // if (
        //   event.shiftKey
        //     ? this.activeCellPos.row - 1 <
        //       Math.min(range.rowStart, range.rowEnd)
        //     : this.activeCellPos.row + 1 >
        //       Math.max(range.rowStart, range.rowEnd)
        // ) {
        //   if (
        //     event.shiftKey
        //       ? this.activeCellPos.column - 1 <
        //         Math.min(range.columnStart, range.columnEnd)
        //       : this.activeCellPos.column + 1 >
        //         Math.max(range.columnStart, range.columnEnd)
        //   ) {
        //     const rangeIndex = event.shiftKey
        //       ? this.activeCellPos.rangeIndex - 1 < 0
        //         ? this.activeArr.length - 1
        //         : this.activeCellPos.rangeIndex - 1
        //       : this.activeCellPos.rangeIndex + 1 > this.activeArr.length - 1
        //       ? 0
        //       : this.activeCellPos.rangeIndex + 1;
        //     range = this.activeArr[rangeIndex];
        //     this.activeCellPos = {
        //       row: event.shiftKey
        //         ? Math.max(range.rowStart, range.rowEnd)
        //         : Math.min(range.rowStart, range.rowEnd),
        //       column: event.shiftKey
        //         ? Math.max(range.columnStart, range.columnEnd)
        //         : Math.min(range.columnStart, range.columnEnd),
        //       rangeIndex,
        //     };
        //   } else {
        //     event.shiftKey
        //       ? this.activeCellPos.column--
        //       : this.activeCellPos.column++;
        //     this.activeCellPos.row = event.shiftKey
        //       ? Math.max(range.rowStart, range.rowEnd)
        //       : Math.min(range.rowStart, range.rowEnd);
        //   }
        // } else {
        //   event.shiftKey ? this.activeCellPos.row-- : this.activeCellPos.row++;
        // }
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
    this.drawRuler(this.ctx);
    this.drawScrollBar(this.ctx);
  }

  onKeyDown(event: KeyboardEvent) {
    console.log('keydown', event);
    switch (event.code) {
      case KeyCode.Tab:
      case KeyCode.Enter:
        event.preventDefault();
        this.onKeyTabOrEnter(event);
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
      case KeyCode.KeyA:
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
    this.drawRuler(this.ctx);
    this.drawScrollBar(this.ctx);
    this.editingCell = null;
    this.canvas.focus();
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

  toggleCombine() {
    const range = this.activeArr[this.activeCellPos.rangeIndex];
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
      for (let i = rowStart; i <= rowEnd; i++) {
        for (let j = columnStart; j <= columnEnd; j++) {
          if (this.cells[i][j].rowSpan > 1 || this.cells[i][j].colSpan > 1) {
            hasCombine = true;
          }
          if (hasCombine) {
            this.cells[i][j].isCombined = false;
            this.cells[i][j].rowSpan = 1;
            this.cells[i][j].colSpan = 1;
            this.cells[i][j].combineCell = null;
          }
        }
      }
      if (!hasCombine) {
        cell.rowSpan = rowEnd - rowStart + 1;
        cell.colSpan = columnEnd - columnStart + 1;
        for (let i = rowStart; i <= rowEnd; i++) {
          for (let j = columnStart; j <= columnEnd; j++) {
            if (i === rowStart && j === columnStart) {
              continue;
            }
            this.cells[i][j].isCombined = true;
            this.cells[i][j].rowSpan = 1;
            this.cells[i][j].colSpan = 1;
            this.cells[i][j].combineCell = cell;
          }
        }
        this.activeArr = [
          {
            rowStart: cell.position.row,
            rowEnd: cell.position.row + cell.rowSpan - 1,
            columnStart: cell.position.column,
            columnEnd: cell.position.column + cell.colSpan - 1,
          },
        ];
      } else {
        this.activeArr = [
          {
            rowStart,
            rowEnd,
            columnStart,
            columnEnd,
          },
        ];
      }

      this.activeCellPos = {
        row: cell.position.row,
        column: cell.position.column,
        rangeIndex: 0,
      };
      this.refreshView();
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
          this.cells[i][j].isCombined = false;
          this.cells[i][j].rowSpan = 1;
          this.cells[i][j].colSpan = 1;
          this.cells[i][j].combineCell = null;
        }
      }
      this.activeArr = [
        {
          rowStart,
          rowEnd,
          columnStart,
          columnEnd,
        },
      ];
      this.activeCellPos = {
        row: rowStart,
        column: columnStart,
        rangeIndex: 0,
      };
      this.refreshView();
    }
    this.canvas.focus();
  }

  recalcRange(range: CellRange, isExpand = true) {
    let rowReverse = range.rowStart > range.rowEnd;
    let columnReverse = range.columnStart > range.columnEnd;
    let needReCalc = false;
    const [rStart, rEnd, cStart, cEnd] = [
      Math.min(range.rowStart, range.rowEnd),
      Math.max(range.rowStart, range.rowEnd),
      Math.min(range.columnStart, range.columnEnd),
      Math.max(range.columnStart, range.columnEnd),
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
        if (this.cells[i][j].hidden) {
          continue;
        }
        const cell = this.cells[i][j].isCombined
          ? this.cells[i][j].combineCell
          : this.cells[i][j];
        if (cell.hidden || (cell.rowSpan === 1 && cell.colSpan === 1)) {
          continue;
        }
        const [isActiveCellRow, isActiveCellColumn] = [
          cell.position.row === this.activeCellPos.row,

          cell.position.column === this.activeCellPos.column,
        ];
        if (cell.position.row < rStart) {
          if (isExpand) {
            rowStart = cell.position.row;
          } else {
            rowStart = isActiveCellRow
              ? cell.position.row
              : cell.position.row + cell.rowSpan;
            rowEnd = isActiveCellRow
              ? cell.position.row + cell.rowSpan
              : rowEnd;
            rowReverse = false;
            needReCalc = true;
          }
          //   rowStart = isExpand ? rowReverse ?
          //   rowEnd = isExpand
          //     ? cell.position.row
          //     : cell.position.row + cell.rowSpan;
          //   rowReverse = false;
        } else if (cell.position.row + cell.rowSpan - 1 > rowEnd) {
          if (isExpand) {
            rowReverse
              ? (rowStart = cell.position.row + cell.rowSpan - 1)
              : (rowEnd = cell.position.row + cell.rowSpan - 1);
          } else {
            rowStart = isActiveCellRow
              ? cell.position.row + cell.rowSpan - 1
              : rowReverse
              ? rowEnd
              : rowStart;
            rowEnd = cell.position.row - 1;
            rowReverse = isActiveCellRow ? false : rowReverse;
          }
          //   rowReverse = false;
          //   rowEnd = isExpand
          //     ? cell.position.row + cell.rowSpan - 1
          //     : cell.position.row - 1;
          //   rowStart = isExpand
          //     ? rowStart
          //     : rowStart === cell.position.row
          //     ? cell.position.row + cell.rowSpan - 1
          //     : rowStart;
        }
        if (cell.position.column < cStart) {
          if (isExpand) {
            columnStart = cell.position.column;
          } else {
            columnStart = isActiveCellColumn
              ? cell.position.column
              : cell.position.column + cell.colSpan;
            columnEnd = isActiveCellColumn
              ? cell.position.column + cell.colSpan
              : columnEnd;
            columnReverse = isActiveCellColumn ? false : columnReverse;
            needReCalc = true;
          }
          //   columnStart = isExpand
          //     ? columnReverse
          //       ? columnStart
          //       : cell.position.column
          //     : cell.position.column;
          //   columnEnd = isExpand
          //     ? cell.position.column
          //     : cell.position.column + cell.colSpan;
          //   columnReverse = false;
        } else if (cell.position.column + cell.colSpan - 1 > cEnd) {
          if (isExpand) {
            columnReverse
              ? (columnStart = cell.position.column + cell.colSpan - 1)
              : (columnEnd = cell.position.column + cell.colSpan - 1);
          } else {
            columnStart = isActiveCellColumn
              ? cell.position.column + cell.colSpan - 1
              : columnReverse
              ? columnEnd
              : columnStart;
            columnEnd = cell.position.column - 1;
            columnReverse = isActiveCellColumn ? false : columnReverse;
          }
          //   columnReverse = false;
          //   columnEnd = isExpand
          //     ? cell.position.column + cell.colSpan - 1
          //     : cell.position.column - 1;
          //   columnStart = isExpand
          //     ? columnStart
          //     : cell.position.column + cell.colSpan - 1;
        }
      }
    }
    console.log(range);
    range.rowStart = rowReverse ? rowEnd : rowStart;
    range.rowEnd = rowReverse ? rowStart : rowEnd;
    range.columnStart = columnReverse ? columnEnd : columnStart;
    range.columnEnd = columnReverse ? columnStart : columnEnd;
    console.log(range);

    // const endCell = this.cells[range.rowEnd][range.columnEnd].isCombined
    //   ? this.cells[range.rowEnd][range.columnEnd].combineCell
    //   : this.cells[range.rowEnd][range.columnEnd];
    // if (
    //   columnReverse ? endCell.endCell.colSpan + endCell.position.column - 1 > range.columnEnd ||
    //   rowReverse? endCell.position.row < range.rowEnd : endCell.rowSpan + endCell.position.row - 1 > range.rowEnd
    // ) {
    //   this.recalcRange(range, true);
    // }
    if (needReCalc) {
      this.recalcRange(range, true);
    }
  }
}
