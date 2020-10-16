import { AutoFocusDirective } from './directive/autofous.directive';
import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { AfterContentInitDirective } from './directive/aftercontentinit';
import { OnDestroyDirective } from './directive/ondestroy';

@NgModule({
  declarations: [AutoFocusDirective, AfterContentInitDirective, OnDestroyDirective],
  imports: [CommonModule],
  exports: [AutoFocusDirective, AfterContentInitDirective, OnDestroyDirective],
})
export class CoreModule {}
