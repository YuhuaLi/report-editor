import {
  AfterContentInit,
  AfterViewInit,
  ChangeDetectionStrategy,
  ChangeDetectorRef,
  Component,
  ElementRef,
  EventEmitter,
  Input,
  OnDestroy,
  OnInit,
  Output,
  ViewChild,
} from '@angular/core';
import { DomSanitizer } from '@angular/platform-browser';
import { Cell } from 'src/app/core/model';
import { OperateState } from '../../../core/model/operate-state.enum';

@Component({
  selector: 'app-cell-edit',
  templateUrl: './cell-edit.component.html',
  styleUrls: ['./cell-edit.component.scss'],
})
export class CellEditComponent implements OnInit, AfterViewInit, OnDestroy {
  @Input() editingCell: Cell;
  @Input() state;
  @Input() multiple = 1;
  @Input() offsetLeft = 0;
  @Input() offsetTop = 0;
  @Input() scrollTop = 0;
  @Input() scrollLeft = 0;
  @Output() onkeydown: EventEmitter<any> = new EventEmitter<any>();
  @Output() afterInit: EventEmitter<any> = new EventEmitter<any>();
  @ViewChild('editArea') editArea: ElementRef;
  html: any;
  editStatus = OperateState.EditCell;

  constructor(
    private sanitizer: DomSanitizer,
    private cdr: ChangeDetectorRef
  ) {}

  ngOnInit(): void {
    this.html = this.safeHtml(
      this.editingCell.content.html || this.editingCell.content.value
    );
  }

  ngAfterViewInit() {
    this.afterInit.emit(this.editArea.nativeElement);
  }

  safeHtml(html) {
    return (html && this.sanitizer.bypassSecurityTrustHtml(html)) || '';
  }

  onpaste(event) {
    console.log('paste', event);
  }

  onInput(event) {
    if (this.editingCell) {
      this.editingCell.content.value = event.target.textContent.replace(
        /&nbsp;/g,
        ' '
      );
    }
  }

  ngOnDestroy() {
    // this.editingCell.content.html = this.editArea.nativeElement.innerHTML;
    // console.log(this.editArea.nativeElement.childNodes);
    // console.log(this.editArea.nativeElement.children);
    // console.log(this.editArea.nativeElement.textContent);
  }
}
