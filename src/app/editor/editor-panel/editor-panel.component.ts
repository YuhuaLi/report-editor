import {
  Component,
  OnInit,
  ViewChild,
  ElementRef,
  AfterViewInit,
  HostListener,
  ChangeDetectorRef,
  ChangeDetectionStrategy,
} from '@angular/core';
import { Panel } from './panel';

@Component({
  selector: 'app-editor-panel',
  templateUrl: './editor-panel.component.html',
  styleUrls: ['./editor-panel.component.scss'],
  changeDetection: ChangeDetectionStrategy.OnPush,
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

  fontFamilyArr = [
    'sans-serif',
    'Arial',
    'SimSun',
    'SimHei',
    'Microsoft YaHei',
    'KaiTi',
    'Microsoft JhengHei',
  ];
  fontSizeArr = Array.from({ length: 30 }).map((val, idx) => idx + 10);

  showFontColorPanel = false;
  showBackColorPanel = false;

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
    if (this.multiple < 1.5) {
      this.multiple += 0.1;
      this.init();
      this.cdr.detectChanges();
    }
  }
  zoomOut() {
    if (this.multiple > 0.5) {
      this.multiple -= 0.1;
      this.init();
      this.cdr.detectChanges();
    }
  }

  onBarInputFocus(event) {
    event.preventDefault();
    if (!this.editingCell) {
      this.editingCell = this.activeCell;
      this.editingCell.content.previousValue = this.editingCell.content.value;
    }
  }

  toggleFontColorPanel() {
    this.showFontColorPanel = !this.showFontColorPanel;
  }
  toggleBackColorPanel() {
    this.showBackColorPanel = !this.showBackColorPanel;
  }
  changeColor(event) {
    if (this.showFontColorPanel) {
      this.changeCellStyle({ color: event.target.style.backgroundColor });
    } else if (this.showBackColorPanel) {
      this.changeCellStyle({ background: event.target.style.backgroundColor });
    }
  }

  insertImage(event) {
    console.log(event);
    if (/image\/\w*/.test(event.target.files[0].type)) {
      createImageBitmap(event.target.files[0]).then((img) => {
        this.addImage(img);
      });
    }
  }
}
