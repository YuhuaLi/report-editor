import { AutoFocusDirective } from './directive/autofous.directive';
import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';

@NgModule({
  declarations: [AutoFocusDirective],
  imports: [CommonModule],
  exports: [AutoFocusDirective],
})
export class CoreModule {}
