import {
  AfterContentInit,
  Directive,
  ElementRef,
  EventEmitter,
  Output,
} from '@angular/core';

@Directive({ selector: '[appAfterContentInit]' })
export class AfterContentInitDirective implements AfterContentInit {
  @Output('appAfterContentInit')
  public after: EventEmitter<void> = new EventEmitter<void>();

  constructor(private elementRef: ElementRef) {}

  public ngAfterContentInit(): void {
    setTimeout(() => this.after.next(this.elementRef.nativeElement));
  }
}
