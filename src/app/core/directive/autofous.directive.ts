import { Directive, OnInit, ElementRef, Input } from '@angular/core';

@Directive({
  selector: '[appAutoFocus]',
})
export class AutoFocusDirective implements OnInit {
  @Input() appAutoFocus: boolean;
  constructor(private elementRef: ElementRef) {}

  ngOnInit(): void {
    if (this.appAutoFocus) {
      this.elementRef.nativeElement.focus();
    }
  }
}
