import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { AutoFocusDirective } from './directive/autofous.directive';

@NgModule({
  declarations: [AutoFocusDirective],
  imports: [CommonModule],
  exports: [AutoFocusDirective]
})
export class CoreModule {}
