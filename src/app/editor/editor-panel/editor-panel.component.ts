import {
  Component,
  OnInit,
  ViewChild,
  ElementRef,
  AfterViewInit,
  HostListener,
  ChangeDetectorRef,
  ChangeDetectionStrategy,
  ViewContainerRef,
} from '@angular/core';
import { DomSanitizer } from '@angular/platform-browser';
import { Panel } from './panel';

@Component({
  selector: 'app-editor-panel',
  templateUrl: './editor-panel.component.html',
  styleUrls: ['./editor-panel.component.scss'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class EditorPanelComponent
  extends Panel
  implements OnInit, AfterViewInit
{
  @ViewChild('panel') panel: ElementRef;
  @ViewChild('actionPanel') actionPanel: ElementRef;
  @ViewChild('animationPanel') animationPanel: ElementRef;
  @ViewChild('floatPanel') floatPanel: ElementRef;
  @ViewChild('floatActionPanel') floatActionPanel: ElementRef;
  @ViewChild('cellEditor') cellEditor: any;
  zoomSize: number;
  selection = [];

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

  constructor(
    private elmentRef: ElementRef,
    private cdr: ChangeDetectorRef,
    private sanitizer: DomSanitizer
  ) {
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

  editCellCompelte(change = true) {
    if (change) {
      this.editingCell.content.html =
        this.cellEditor.editArea.nativeElement.innerHTML;
    }
    super.editCellCompelte(change);
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

  onBarInputMouseup(event) {
    // console.log(event.target.selectionStart, event.target.selectionEnd);
    const input = event.target;
    if (input.selectionStart !== input.selectionEnd) {
      this.selection = [input.selectionStart, input.selectionEnd];
    } else {
      this.selection = [];
    }
  }

  onBarInputKeypress(event) {
    // console.log(event.target.selectionStart, event.target.selectionEnd);
    const input = event.target;
    if (input.selectionStart !== input.selectionEnd) {
      this.selection = [input.selectionStart, input.selectionEnd];
    } else {
      this.selection = [];
    }
  }

  onBarInputFocus(event) {
    // event.preventDefault();
    event.target.focus();
    if (!this.editingCell) {
      this.editingCell = this.activeCell;
      this.editingCell.content.previousValue = this.editingCell.content.value;
      this.editingCell.content.previousHtml = this.editingCell.content.html;
    }
    // console.log(this.getSelection(), this.getSelection().toString());
  }

  toggleFontColorPanel(event) {
    // event.preventDefault();
    this.showFontColorPanel = !this.showFontColorPanel;
    // this.state.isCellEdit = true;
  }
  toggleBackColorPanel() {
    this.showBackColorPanel = !this.showBackColorPanel;
    // this.state.isCellEdit = true;
  }
  changeColor(event) {
    if (this.showFontColorPanel) {
      this.changeStyle({ color: event.target.style.backgroundColor });
    } else if (this.showBackColorPanel) {
      this.changeStyle({ background: event.target.style.backgroundColor });
    }
    this.cdr.detectChanges();
  }

  insertImage(event) {
    for (const file of event.target.files) {
      if (/image\/\w*/.test(file.type)) {
        createImageBitmap(event.target.files[0]).then((img) => {
          this.addImage(img);
          event.target.value = null;
        });
        break;
      }
    }
    this.canvas.focus();
  }

  getSelection() {
    return window.getSelection
      ? window.getSelection()
      : document.getSelection
      ? document.getSelection()
      : null;
  }

  changeStyle(style) {
    const selection = this.getSelection();
    if (selection && selection.toString()) {
      document.execCommand('styleWithCss', false);
      Object.keys(style).forEach((key) => {
        switch (key) {
          case 'fontWeight':
            document.execCommand('bold', false);
            break;
          case 'color':
            document.execCommand('foreColor', false, style[key]);
            break;
          case 'fontSize':
            document.execCommand('fontSize', false, style[key]);
            selection.anchorNode.parentElement.style.fontSize =
              style[key] + 'pt';
            break;
          case 'fontFamily':
            document.execCommand('fontName', false, style[key]);
            break;
          case 'fontStyle':
            document.execCommand('italic', false);
            break;
          default:
            this.changeCellStyle(style);
            break;
        }
      });
    } else {
      this.changeCellStyle(style);
    }
  }

  afterEditCellInit(event) {
    const selection = this.getSelection();
    if (selection && selection.anchorOffset === 0) {
      const range = document.createRange();
      range.selectNodeContents(event);
      range.collapse(false);
      selection.removeAllRanges();
      selection.addRange(range);
    }
  }

  clickWithoutFocus(event) {
    const selection = this.getSelection();
    event.preventDefault();
    return false;
  }

  onEditCellDestory(event) {
    this.editingCell.content.html = this.sanitizer.bypassSecurityTrustHtml(
      event.target.innerHTML
    );
  }

  deleteCellValue(prev, cur, element, index) {
    const textObj = this.parseNode(element);
    let del;
    if (cur.length === index) {
      del = prev.substring(index);
    } else {
      del = prev.substr(index, prev.length - cur.length);
    }
    let curIndex = 0;
    for (const obj of textObj) {
      const textLen = obj.text.length;
      // if (curIndex <= index && curIndex + obj.text.length > index) {
      //   if (del.length < obj.text.length) {
      //     obj.node.textContent = obj.node.textContent.replace(del, '');
      //   } else {
      //     if (curIndex === index) {
      //       obj.node.parentNode.removeChild(obj.node);
      //     } else {
      //       obj.node.textContent = obj.node.textContent.substring(
      //         curIndex,
      //         index
      //       );
      //     }
      //     del = del.substring(textLen - index + curIndex);
      //     index += textLen - index + curIndex;
      //   }
      // }
      if (curIndex <= index && curIndex + obj.text.length > index) {
        if (curIndex === index && obj.text.length <= del.length) {
          obj.node.parentNode.removeChild(obj.node);
        } else {
          obj.node.textContent =
            obj.node.textContent.substring(0, index - curIndex) +
            obj.node.textContent.substring(index - curIndex + del.length);
        }
        del = del.substring(textLen - index + curIndex);
        index += textLen - index + curIndex;
      }
      curIndex += textLen;
    }
  }

  refreshEditCellValue(event, input) {
    let prev = this.editingCell.content.value || '';
    const cur = event;

    // const element = this.htmlToElement(
    //   this.editingCell.content.html || this.editingCell.content.value
    // );
    const element = this.cellEditor.editArea.nativeElement;

    if (this.selection.length) {
      this.deleteCellValue(
        prev,
        prev.substring(0, this.selection[0]) + prev.substring(this.selection[1]),
        element,
        this.selection[0]
      );
      // input.setSelectionRange(
      //   cur.length - prev.length + this.selection[1],
      //   cur.length - prev.length + this.selection[1]
      // );
      prev = prev.replace(
        prev.substring(this.selection[0], this.selection[1]),
        ''
      );
      this.selection = [];
    }
    const textObj = this.parseNode(element);
    if (prev.length === cur.length) {
      // replace
      let i = 0;
      let index = -1;
      let diffLen = 0;
      while (i < cur.length) {
        if (cur[i] !== prev[i] && index === -1) {
          index = i;
        }
        if (cur.substring(i) === prev.substring(i) && index > -1) {
          break;
        }
        if (index > -1) {
          diffLen++;
        }
        i++;
      }
      let diff = cur.substring(index, index + diffLen);
      let curIndex = 0;
      for (const obj of textObj) {
        if (curIndex <= index && curIndex + obj.text.length > index) {
          if (diff.length <= obj.text.length) {
            const charArr = obj.node.textContent.split('');
            charArr.splice(index - curIndex, diff.length, ...diff);
            obj.node.textContent = charArr.join('');
            break;
          } else {
            const charArr = obj.node.textContent.split('');
            charArr.splice(
              index - curIndex,
              obj.text.length,
              ...diff.substring(0, obj.text.length - index + curIndex)
            );
            obj.node.textContent = charArr.join('');
            diff = diff.substring(obj.text.length - index + curIndex);
            index += obj.text.length - index + curIndex;
          }
        }
        curIndex += obj.text.length;
      }
      // this.cellEditor.cdr.detectChanges();
      // this.cdr.detectChanges();
    } else if (prev.length > cur.length) {
      // delete
      let index = input.selectionStart;
      let del;
      if (cur.length === index) {
        del = prev.substring(index);
      } else {
        del = prev.substr(index, prev.length - cur.length);
      }
      let curIndex = 0;
      for (const obj of textObj) {
        const textLen = obj.text.length;
        if (curIndex <= index && curIndex + obj.text.length > index) {
          if (del.length < obj.text.length) {
            obj.node.textContent =
              obj.node.textContent.substring(0, index - curIndex) +
              obj.node.textContent.substring(index - curIndex + del.length);
            // obj.node.textContent = obj.node.textContent.replace(del, '');
          } else {
            if (curIndex === index) {
              obj.node.parentNode.removeChild(obj.node);
            } else {
              obj.node.textContent = obj.node.textContent.substring(
                curIndex,
                index
              );
            }
            del = del.substring(textLen - index + curIndex);
            index += textLen - index + curIndex;
          }
        }
        curIndex += textLen;
      }
    } else {
      // add
      const index = input.selectionStart - cur.length + prev.length;
      const add = cur.substr(index, cur.length - prev.length);
      let curIndex = 0;
      for (const obj of textObj) {
        const textLen = obj.text.length;
        if (curIndex <= index && curIndex + obj.text.length > index) {
          obj.node.textContent =
            obj.node.textContent.substring(0, index - curIndex) +
            add +
            obj.node.textContent.substring(index - curIndex);
          break;
        }
        curIndex += textLen;
      }
      if (input.selectionStart === cur.length && textObj.length) {
        textObj[textObj.length - 1].node.textContent += add;
      } else if (!textObj.length) {
        element.textContent = add;
      }
    }

    // if (!this.editingCell) {
    //   this.cellEditor.editArea.nativeElement.innerHTML = '';
    //   this.cellEditor.editArea.nativeElement.appendChild(element);
    //   this.editingCell.content.html = this.cellEditor.editArea.nativeElement.innerHTML;
    // }
    this.editingCell.content.value = event;

    // this.cellEditor.editArea.nativeElement.textContent = event;
  }
}
