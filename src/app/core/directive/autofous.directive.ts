import { Directive, OnInit, ElementRef, Input, OnChanges, SimpleChanges } from '@angular/core';

@Directive({
  selector: '[appAutoFocus]',
})
export class AutoFocusDirective implements OnInit, OnChanges {
  @Input() appAutoFocus: boolean;
  constructor(private elementRef: ElementRef) {}

  ngOnInit(): void {
    if (this.appAutoFocus) {
      this.elementRef.nativeElement.focus();
    }
  }
  ngOnChanges(changes: SimpleChanges) {
    if (this.appAutoFocus) {
      this.elementRef.nativeElement.focus();
    }
  }
}
