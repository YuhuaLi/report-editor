import {
  Directive,
  ElementRef,
  EventEmitter,
  OnDestroy,
  Output,
} from '@angular/core';

@Directive({ selector: '[appOnDestroy]' })
export class OnDestroyDirective implements OnDestroy {
  @Output('appOnDestroy')
  public destory: EventEmitter<void> = new EventEmitter<void>();

  constructor(private elementRef: ElementRef) {}

  public ngOnDestroy(): void {
    setTimeout(() => {this.destory.next(this.elementRef.nativeElement);
    console.log('destroy')});
  }
}
