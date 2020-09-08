import {
  Component,
  OnInit,
  ViewChild,
  ElementRef,
  AfterViewInit,
  HostListener,
  ChangeDetectorRef,
} from '@angular/core';
import { Panel } from './panel';

@Component({
  selector: 'app-editor-panel',
  templateUrl: './editor-panel.component.html',
  styleUrls: ['./editor-panel.component.scss'],
})
export class EditorPanelComponent
  extends Panel
  implements OnInit, AfterViewInit {
  @ViewChild('panel') panel: ElementRef;
  @ViewChild('actionPanel') actionPanel: ElementRef;
  @ViewChild('animationPanel') animationPanel: ElementRef;
  @ViewChild('floatPanel') floatPanel: ElementRef;
  @ViewChild('floatActionPanel') floatActionPanel: ElementRef;
  zoomSize: number;
  // width = 0;
  // height = 0;
  // viewRowCount = 0;
  // viewColumnCount = 0;
  // offsetLeft = 0;
  // offsetTop = 30;
  // columns = [];
  // rows = [];
  // cells: Cell[][] = [];
  // viewCells: Cell[][] = [];
  // editingCell: Cell;
  // scrollBarWidth = Style.scrollBarWidth;
  // // activeRange: {
  // //   rowStart: number;
  // //   columnStart: number;
  // //   rowEnd: number;
  // //   columnEnd: number;
  // // } = { rowStart: 1, columnStart: 1, rowEnd: 1, columnEnd: 1 };
  // state: any = {
  //   isSelectCell: false,
  //   isScrollYThumbHover: false,
  //   isSelectScrollYThumb: false,
  // };
  // isTicking = false;
  // scrollLeft = 0;
  // scrollTop = 0;
  // scrollWidth = 0;
  // scrollHeight = 0;
  // clientWidth = 0;
  // clientHeight = 0;
  // mousePoint: any;
  // autoScrollTimeoutID: any;
  // activeCellPos: any = { row: 1, column: 1, rangeIndex: 0 };
  // activeArr: CellRange[] = [
  //   { rowStart: 1, columnStart: 1, rowEnd: 1, columnEnd: 1 },
  // ];
  // unActiveRange: CellRange;
  // resizeColumnCell: Cell;
  // resizeRowCell: Cell;

  // ctx: CanvasRenderingContext2D;
  // actionCtx: CanvasRenderingContext2D;

  constructor(private elmentRef: ElementRef, private cdr: ChangeDetectorRef) {
    super();
  }

  @HostListener('window:resize', ['$event'])
  onResize(event) {
    console.log('resize');
    this.resize();
    this.refreshView();
  }

  ngOnInit(): void {}

  ngAfterViewInit() {
    this.canvas = this.panel.nativeElement;
    this.actionCanvas = this.actionPanel.nativeElement;
    this.animationCanvas = this.animationPanel.nativeElement;
    this.floatCanvas = this.floatPanel.nativeElement;
    this.floatActionCanvas = this.floatActionPanel.nativeElement;
    this.init();
    this.cdr.detectChanges();
  }

  zoomIn() {
    this.multiple = this.multiple + 0.2 > 2 ? 2 : this.multiple + 0.2;
    this.init();
    this.cdr.detectChanges();
  }
  zoomOut() {
    this.multiple = this.multiple - 0.2 < 0.2 ? 0.2 : this.multiple - 0.2;
    this.init();
    this.cdr.detectChanges();
  }
}
